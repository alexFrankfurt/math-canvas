#include "math_manager.h"
#include "math_evaluator.h"
#include <algorithm>
#include <cmath>

void MathManager::ShiftObjectsAfter(LONG atPosInclusive, LONG delta)
{
    if (delta == 0) return;
    for (auto it = m_objects.begin(); it != m_objects.end(); )
    {
        if (it->barStart >= atPosInclusive)
        {
            it->barStart += delta;
            if (it->barStart < 0)
            {
                it = m_objects.erase(it);
                continue;
            }
        }
        ++it;
    }
}

void MathManager::DeleteObjectsInRange(LONG start, LONG end)
{
    if (start == end) return;
    if (start > end) std::swap(start, end);

    for (ptrdiff_t i = (ptrdiff_t)m_objects.size() - 1; i >= 0; --i)
    {
        auto& obj = m_objects[i];
        LONG objEnd = obj.barStart + obj.barLen;
        if (!(end <= obj.barStart || start >= objEnd))
        {
            m_objects.erase(m_objects.begin() + i);
        }
    }
}

bool MathManager::IsPosInsideAnyObject(LONG pos, size_t* outIndex)
{
    for (size_t i = 0; i < m_objects.size(); ++i)
    {
        const auto& obj = m_objects[i];
        if (pos >= obj.barStart && pos < (obj.barStart + obj.barLen))
        {
            if (outIndex) *outIndex = i;
            return true;
        }
    }
    return false;
}

double MathManager::CalculateResult(const MathObject& obj)
{
    MathEvaluator eval;
    if (obj.type == MathType::Fraction)
    {
        if (obj.part2.empty()) return 0;
        double den = _wtof(obj.part2.c_str());
        if (den == 0) return 0;
        return _wtof(obj.part1.c_str()) / den;
    }
    else if (obj.type == MathType::Summation)
    {
        std::wstring var; double start;
        if (!ParseLowerLimit(obj.part2, var, start)) return 0;
        double end = _wtof(obj.part1.c_str());
        double sum = 0;
        for (double i = start; i <= end; ++i)
            sum += eval.Eval(obj.part3, var, i);
        return sum;
    }
    else if (obj.type == MathType::SystemOfEquations)
    {
        // No single numeric result for a system of equations
        return 0;
    }
    else if (obj.type == MathType::SquareRoot)
    {
        double val = eval.Eval(obj.part1);
        if (obj.part2.empty() || obj.part2 == L"2")
            return sqrt(val);
        double n = _wtof(obj.part2.c_str());
        if (n == 0) return 0;
        return pow(val, 1.0 / n);
    }
    else if (obj.type == MathType::Integral)
    {
        double a = _wtof(obj.part2.c_str());
        double b = _wtof(obj.part1.c_str());
        std::wstring var = L"x";
        size_t dPos = obj.part3.find(L" d");
        std::wstring expr = obj.part3;
        if (dPos != std::wstring::npos)
        {
            if (dPos + 2 < obj.part3.size()) var = obj.part3.substr(dPos + 2, 1);
            expr = obj.part3.substr(0, dPos);
        }
        int steps = 200;
        double dx = (b - a) / steps;
        double sum = 0;
        for (int i = 0; i <= steps; ++i)
        {
            double x = a + i * dx;
            double fx = eval.Eval(expr, var, x);
            if (i == 0 || i == steps) sum += fx / 2.0;
            else sum += fx;
        }
        return sum * dx;
    }
    return 0;
}
