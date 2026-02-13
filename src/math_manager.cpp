#include "math_manager.h"
#include "math_evaluator.h"
#include <algorithm>
#include <cmath>
#include <sstream>
#include <iomanip>
#include <iostream>

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
        // System of equations are handled by CalculateSystemResult method
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
    else if (obj.type == MathType::AbsoluteValue)
    {
        double val = eval.Eval(obj.part1);
        return fabs(val);
    }
    else if (obj.type == MathType::Power)
    {
        double base = eval.Eval(obj.part1);
        double exponent = eval.Eval(obj.part2);
        return pow(base, exponent);
    }
    else if (obj.type == MathType::Logarithm)
    {
        double arg = eval.Eval(obj.part2);
        if (arg <= 0) return 0;  // Invalid argument
        
        // part1 is the base (default 10)
        double base = 10.0;
        if (!obj.part1.empty()) {
            base = eval.Eval(obj.part1);
        }
        if (base <= 0 || base == 1) return 0;  // Invalid base
        
        return log(arg) / log(base);
    }
    return 0;
}

std::wstring MathManager::CalculateSystemResult(const MathObject& obj)
{
    MathEvaluator eval;
    std::vector<std::wstring> equations;

    // Parse equations from the parts - for system of equations, we use part1 and part2 as separate equations
    // and part3 could be a third equation if present
    if (!obj.part1.empty()) equations.push_back(obj.part1);
    if (!obj.part2.empty()) equations.push_back(obj.part2);
    if (!obj.part3.empty()) equations.push_back(obj.part3);

    // Solve the system using rational arithmetic for exact results
    auto solution = eval.SolveSystemOfEquationsRational(equations);

    // Format the result
    long long status = 0;
    if (solution.find(L"status") != solution.end()) {
        status = solution[L"status"].num;  // Get numerator of status rational
    }

    if (status == 0) {
        // Build result string directly
        std::wstring result = L" \uFF1D ";  // Full-width equals sign
        bool first = true;
        for (const auto& [var, val] : solution) {
            if (var == L"status") continue;

            if (!first) {
                result += L", ";
            }

            result += var + L"=";
            
            // Format as fraction if denominator is not 1, otherwise as integer
            if (val.den == 1) {
                result += std::to_wstring(val.num);
            } else {
                result += val.toString();  // Will format as "num/den"
            }

            first = false;
        }

        // If no variables were added (all zero), explicitly add them
        if (first) {
            result += L"x=0, y=0";
        }
        return result;
    } else {
        // Error cases
        switch (status) {
            case -1: return L" \uFF1D Infinite solutions";
            case -2: return L" \uFF1D No solution";
            case -3: return L" \uFF1D No equations";
            case -4: return L" \uFF1D Parse error";
            case -5: return L" \uFF1D Underdetermined system";
            case -6: return L" \uFF1D Too many equations (max 3)";
            default: return L" \uFF1D Unknown error";
        }
    }
}
