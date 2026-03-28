#include "math_manager.h"
#include "math_evaluator.h"
#include <algorithm>
#include <cmath>
#include <sstream>
#include <iomanip>
#include <iostream>

namespace
{
    static std::wstring FormatMessageResult(const std::wstring& message)
    {
        return L" \uFF1D " + message;
    }

    static std::wstring FormatBareNumber(double value)
    {
        if (!std::isfinite(value))
            return L"undefined";
        if (std::fabs(value) < 1e-12)
            value = 0;

        std::wostringstream stream;
        const double nearestInt = std::round(value);
        if (std::fabs(value - nearestInt) < 1e-9)
        {
            stream << (long long)nearestInt;
            return stream.str();
        }

        stream << std::fixed << std::setprecision(6) << value;
        std::wstring text = stream.str();
        const size_t dot = text.find(L'.');
        if (dot != std::wstring::npos)
        {
            size_t end = text.size();
            while (end > dot + 1 && text[end - 1] == L'0')
                --end;
            if (end > dot && text[end - 1] == L'.')
                --end;
            text.resize(end);
        }
        return text;
    }

    static std::wstring TrimCopy(const std::wstring& text)
    {
        const size_t first = text.find_first_not_of(L" \t");
        if (first == std::wstring::npos) return L"";
        const size_t last = text.find_last_not_of(L" \t");
        return text.substr(first, last - first + 1);
    }

    static void AddDimension(UnitDimension& target, const UnitDimension& source, int sign)
    {
        target.length += source.length * sign;
        target.mass += source.mass * sign;
        target.time += source.time * sign;
        target.current += source.current * sign;
        target.temperature += source.temperature * sign;
        target.amount += source.amount * sign;
        target.luminousIntensity += source.luminousIntensity * sign;
    }

    static MathValue NormalizeDisplay(MathValue value)
    {
        if (value.IsError())
            return value;

        if (value.IsDimensionless())
        {
            value.displayScale = 1.0;
            value.displayUnit.clear();
            return value;
        }

        if (value.displayUnit.empty())
        {
            value.displayUnit = BuildCanonicalUnitSymbol(value.dimension);
            value.displayScale = 1.0;
        }
        return value;
    }

    static MathValue AddAccumulatedValues(const MathValue& left, const MathValue& right)
    {
        if (left.IsError()) return left;
        if (right.IsError()) return right;
        if (left.dimension != right.dimension)
            return MathValue::Error(L"incompatible units");

        MathValue result = MathValue::Scalar(left.baseValue + right.baseValue);
        result.dimension = left.dimension;
        if (!left.IsDimensionless())
        {
            if (left.HasDisplayUnit())
            {
                result.displayUnit = left.displayUnit;
                result.displayScale = left.displayScale;
            }
            else if (right.HasDisplayUnit())
            {
                result.displayUnit = right.displayUnit;
                result.displayScale = right.displayScale;
            }
        }

        return NormalizeDisplay(result);
    }

    static MathValue MultiplyAccumulatedValues(const MathValue& left, const MathValue& right)
    {
        if (left.IsError()) return left;
        if (right.IsError()) return right;

        MathValue result = MathValue::Scalar(left.baseValue * right.baseValue);
        result.dimension = left.dimension;
        AddDimension(result.dimension, right.dimension, 1);
        return NormalizeDisplay(result);
    }

    static bool ParseMatrixRows(const MathObject& obj, std::vector<std::vector<double>>& matrix)
    {
        matrix.clear();
        MathEvaluator eval;

        if (obj.slots.size() >= 4)
        {
            std::vector<double> firstRow;
            std::vector<double> secondRow;
            const std::wstring cells[] = {
                TrimCopy(obj.SlotText(1)), TrimCopy(obj.SlotText(2)),
                TrimCopy(obj.SlotText(3)), TrimCopy(obj.SlotText(4))
            };

            for (size_t cellIndex = 0; cellIndex < 4; ++cellIndex)
            {
                if (cells[cellIndex].empty())
                    return false;

                const double value = eval.Eval(cells[cellIndex]);
                if (cellIndex < 2)
                    firstRow.push_back(value);
                else
                    secondRow.push_back(value);
            }

            matrix.push_back(std::move(firstRow));
            matrix.push_back(std::move(secondRow));
            return true;
        }

        const std::wstring rows[] = { obj.SlotText(1), obj.SlotText(2), obj.SlotText(3) };
        size_t expectedCols = 0;

        for (const auto& rowTextRaw : rows)
        {
            const std::wstring rowText = TrimCopy(rowTextRaw);
            if (rowText.empty()) continue;

            std::vector<double> row;
            size_t start = 0;
            while (start <= rowText.size())
            {
                size_t comma = rowText.find(L',', start);
                std::wstring cell = TrimCopy(rowText.substr(start, comma == std::wstring::npos ? std::wstring::npos : comma - start));
                if (cell.empty()) return false;
                row.push_back(eval.Eval(cell));
                if (comma == std::wstring::npos) break;
                start = comma + 1;
            }

            if (row.empty()) return false;
            if (expectedCols == 0) expectedCols = row.size();
            else if (row.size() != expectedCols) return false;
            matrix.push_back(std::move(row));
        }

        return !matrix.empty();
    }
}

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

bool MathManager::CanCalculateResult(const MathObject& obj) const
{
    return obj.type != MathType::Matrix && obj.type != MathType::SystemOfEquations;
}

std::wstring MathManager::CalculateFormattedResult(const MathObject& obj) const
{
    if (obj.type == MathType::Determinant)
    {
        std::vector<std::vector<double>> matrix;
        if (!ParseMatrixRows(obj, matrix))
            return FormatMessageResult(L"invalid matrix");
        if (!(matrix.size() == 2 || matrix.size() == 3))
            return FormatMessageResult(L"unsupported matrix size");
        if (matrix[0].size() != matrix.size())
            return FormatMessageResult(L"matrix must be square");
        return FormatNumericResult(CalculateResult(obj));
    }

    return FormatValueResult(CalculateValueResult(obj));
}

std::wstring MathManager::FormatNumericResult(double value) const
{
    return L" \uFF1D " + FormatBareNumber(value);
}

std::wstring MathManager::FormatValueResult(const MathValue& value) const
{
    if (value.IsError())
        return FormatMessageResult(value.errorText);
    if (!std::isfinite(value.baseValue))
        return FormatMessageResult(L"undefined");
    if (value.IsDimensionless())
        return FormatNumericResult(value.baseValue);

    const double displayScale = std::fabs(value.displayScale) < 1e-12 ? 1.0 : value.displayScale;
    std::wstring result = L" \uFF1D " + FormatBareNumber(value.baseValue / displayScale);
    if (!value.displayUnit.empty())
        result += L" " + value.displayUnit;
    return result;
}

MathValue MathManager::CalculateValueResult(const MathObject& obj) const
{
    MathEvaluator eval;

    if (obj.type == MathType::Fraction)
    {
        const std::wstring numerator = TrimCopy(obj.SlotText(1));
        const std::wstring denominator = TrimCopy(obj.SlotText(2));
        if (numerator.empty() || denominator.empty())
            return MathValue::Error(L"incomplete");
        return eval.EvalValue(L"((" + numerator + L")/(" + denominator + L"))");
    }

    if (obj.type == MathType::Summation)
    {
        const std::wstring upperText = TrimCopy(obj.SlotText(1));
        const std::wstring lowerText = TrimCopy(obj.SlotText(2));
        const std::wstring bodyText = TrimCopy(obj.SlotText(3));
        if (upperText.empty() || lowerText.empty() || bodyText.empty())
            return MathValue::Error(L"incomplete");

        std::wstring var;
        double start = 0;
        if (!ParseLowerLimit(lowerText, var, start))
            return MathValue::Error(L"invalid limits");

        const MathValue upperValue = eval.EvalValue(upperText);
        if (upperValue.IsError())
            return upperValue;
        if (!upperValue.IsDimensionless())
            return MathValue::Error(L"invalid limits");

        MathValue sum = MathValue::Scalar(0.0);
        bool hasTerm = false;
        for (double i = start; i <= upperValue.baseValue; ++i)
        {
            MathValue termValue = eval.EvalValue(bodyText, var, MathValue::Scalar(i));
            if (termValue.IsError())
                return termValue;

            if (!hasTerm)
            {
                sum = termValue;
                hasTerm = true;
            }
            else
            {
                sum = AddAccumulatedValues(sum, termValue);
                if (sum.IsError())
                    return sum;
            }
        }
        return hasTerm ? NormalizeDisplay(sum) : MathValue::Scalar(0.0);
    }

    if (obj.type == MathType::Product)
    {
        const std::wstring upperText = TrimCopy(obj.SlotText(1));
        const std::wstring lowerText = TrimCopy(obj.SlotText(2));
        const std::wstring bodyText = TrimCopy(obj.SlotText(3));
        if (upperText.empty() || lowerText.empty() || bodyText.empty())
            return MathValue::Error(L"incomplete");

        std::wstring var;
        double start = 0;
        if (!ParseLowerLimit(lowerText, var, start))
            return MathValue::Error(L"invalid limits");

        const MathValue upperValue = eval.EvalValue(upperText);
        if (upperValue.IsError())
            return upperValue;
        if (!upperValue.IsDimensionless())
            return MathValue::Error(L"invalid limits");

        MathValue product = MathValue::Scalar(1.0);
        for (double i = start; i <= upperValue.baseValue; ++i)
        {
            product = MultiplyAccumulatedValues(product, eval.EvalValue(bodyText, var, MathValue::Scalar(i)));
            if (product.IsError())
                return product;
        }
        return NormalizeDisplay(product);
    }

    if (obj.type == MathType::Sum)
    {
        const std::wstring expression = TrimCopy(obj.SlotText(1));
        if (expression.empty())
            return MathValue::Error(L"incomplete");
        return eval.EvalValue(expression);
    }

    if (obj.type == MathType::SquareRoot)
    {
        const std::wstring radicand = TrimCopy(obj.SlotText(1));
        const std::wstring indexText = TrimCopy(obj.SlotText(2));
        if (radicand.empty())
            return MathValue::Error(L"incomplete");

        if (indexText.empty() || indexText == L"2")
            return eval.EvalValue(L"sqrt(" + radicand + L")");

        const MathValue indexValue = eval.EvalValue(indexText);
        if (indexValue.IsError())
            return indexValue;
        if (!indexValue.IsDimensionless() || std::fabs(indexValue.baseValue) < 1e-12)
            return MathValue::Error(L"invalid index");

        return eval.EvalValue(L"((" + radicand + L")^(1/(" + indexText + L")))");
    }

    if (obj.type == MathType::Integral)
    {
        const std::wstring upperText = TrimCopy(obj.SlotText(1));
        const std::wstring lowerText = TrimCopy(obj.SlotText(2));
        const std::wstring slotText = TrimCopy(obj.SlotText(3));
        if (upperText.empty() || lowerText.empty() || slotText.empty())
            return MathValue::Error(L"incomplete");

        const MathValue lowerValue = eval.EvalValue(lowerText);
        const MathValue upperValue = eval.EvalValue(upperText);
        if (lowerValue.IsError()) return lowerValue;
        if (upperValue.IsError()) return upperValue;
        if (!lowerValue.IsDimensionless() || !upperValue.IsDimensionless())
            return MathValue::Error(L"invalid limits");

        std::wstring var = L"x";
        std::wstring exprText = slotText;
        const size_t dPos = slotText.find(L" d");
        if (dPos != std::wstring::npos)
        {
            if (dPos + 2 < slotText.size())
                var = slotText.substr(dPos + 2, 1);
            exprText = slotText.substr(0, dPos);
        }

        const int steps = 200;
        const double dx = (upperValue.baseValue - lowerValue.baseValue) / steps;
        MathValue integral = MathValue::Scalar(0.0);
        bool hasSample = false;
        for (int i = 0; i <= steps; ++i)
        {
            const double x = lowerValue.baseValue + i * dx;
            MathValue sample = eval.EvalValue(exprText, var, MathValue::Scalar(x));
            if (sample.IsError())
                return sample;

            sample.baseValue *= (i == 0 || i == steps) ? 0.5 : 1.0;
            if (!hasSample)
            {
                integral = sample;
                hasSample = true;
            }
            else
            {
                integral = AddAccumulatedValues(integral, sample);
                if (integral.IsError())
                    return integral;
            }
        }

        if (!hasSample)
            return MathValue::Scalar(0.0);
        integral.baseValue *= dx;
        return NormalizeDisplay(integral);
    }

    if (obj.type == MathType::AbsoluteValue)
    {
        const std::wstring expression = TrimCopy(obj.SlotText(1));
        if (expression.empty())
            return MathValue::Error(L"incomplete");
        return eval.EvalValue(L"abs(" + expression + L")");
    }

    if (obj.type == MathType::Power)
    {
        const std::wstring base = TrimCopy(obj.SlotText(1));
        const std::wstring exponent = TrimCopy(obj.SlotText(2));
        if (base.empty() || exponent.empty())
            return MathValue::Error(L"incomplete");
        return eval.EvalValue(L"((" + base + L")^(" + exponent + L"))");
    }

    if (obj.type == MathType::Logarithm)
    {
        const std::wstring argText = TrimCopy(obj.SlotText(2));
        if (argText.empty())
            return MathValue::Error(L"incomplete");

        const std::wstring baseText = TrimCopy(obj.SlotText(1));
        if (baseText.empty())
            return eval.EvalValue(L"log(" + argText + L")");
        return eval.EvalValue(L"log_{" + baseText + L"}(" + argText + L")");
    }

    if (obj.type == MathType::Determinant)
        return MathValue::Scalar(CalculateResult(obj));

    return MathValue::Scalar(CalculateResult(obj));
}

double MathManager::CalculateResult(const MathObject& obj) const
{
    MathEvaluator eval;
    if (obj.type == MathType::Fraction)
    {
        const std::wstring& numerator = obj.SlotText(1);
        const std::wstring& denominator = obj.SlotText(2);
        if (denominator.empty()) return 0;
        double den = eval.Eval(denominator);
        if (den == 0) return 0;
        return eval.Eval(numerator) / den;
    }
    else if (obj.type == MathType::Summation)
    {
        std::wstring var; double start;
        if (!ParseLowerLimit(obj.SlotText(2), var, start)) return 0;
        double end = eval.Eval(obj.SlotText(1));
        double sum = 0;
        for (double i = start; i <= end; ++i)
            sum += eval.Eval(obj.SlotText(3), var, i);
        return sum;
    }
    else if (obj.type == MathType::Product)
    {
        // Product of numbers from lower to upper limit (e.g., ∏_{i=a}^{b} expr)
        std::wstring var; double start;
        if (!ParseLowerLimit(obj.SlotText(2), var, start)) return 0;
        double end = eval.Eval(obj.SlotText(1));
        double prod = 1;
        for (double i = start; i <= end; ++i)
            prod *= eval.Eval(obj.SlotText(3), var, i);
        return prod;
    }
    else if (obj.type == MathType::Sum)
    {
        // Sum of numbers separated by + (e.g., "8384 + 384843 + 9138")
        return eval.Eval(obj.SlotText(1));
    }
    else if (obj.type == MathType::SystemOfEquations)
    {
        // System of equations are handled by CalculateSystemResult method
        return 0;
    }
    else if (obj.type == MathType::SquareRoot)
    {
        double val = eval.Eval(obj.SlotText(1));
        if (obj.SlotText(2).empty() || obj.SlotText(2) == L"2")
            return sqrt(val);
        double n = eval.Eval(obj.SlotText(2));
        if (n == 0) return 0;
        return pow(val, 1.0 / n);
    }
    else if (obj.type == MathType::Integral)
    {
        double a = eval.Eval(obj.SlotText(2));
        double b = eval.Eval(obj.SlotText(1));
        std::wstring var = L"x";
        size_t dPos = obj.SlotText(3).find(L" d");
        std::wstring expr = obj.SlotText(3);
        if (dPos != std::wstring::npos)
        {
            if (dPos + 2 < obj.SlotText(3).size()) var = obj.SlotText(3).substr(dPos + 2, 1);
            expr = obj.SlotText(3).substr(0, dPos);
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
        double val = eval.Eval(obj.SlotText(1));
        return fabs(val);
    }
    else if (obj.type == MathType::Power)
    {
        double base = eval.Eval(obj.SlotText(1));
        double exponent = eval.Eval(obj.SlotText(2));
        return pow(base, exponent);
    }
    else if (obj.type == MathType::Logarithm)
    {
        double arg = eval.Eval(obj.SlotText(2));
        if (arg <= 0) return 0;  // Invalid argument
        
        // part1 is the base (default 10)
        double base = 10.0;
        if (!obj.SlotText(1).empty()) {
            base = eval.Eval(obj.SlotText(1));
        }
        if (base <= 0 || base == 1) return 0;  // Invalid base
        
        return log(arg) / log(base);
    }
    else if (obj.type == MathType::Determinant)
    {
        std::vector<std::vector<double>> matrix;
        if (!ParseMatrixRows(obj, matrix)) return 0;
        if (matrix.size() == 2 && matrix[0].size() == 2)
        {
            return matrix[0][0] * matrix[1][1] - matrix[0][1] * matrix[1][0];
        }
        if (matrix.size() == 3 && matrix[0].size() == 3)
        {
            return matrix[0][0] * (matrix[1][1] * matrix[2][2] - matrix[1][2] * matrix[2][1])
                - matrix[0][1] * (matrix[1][0] * matrix[2][2] - matrix[1][2] * matrix[2][0])
                + matrix[0][2] * (matrix[1][0] * matrix[2][1] - matrix[1][1] * matrix[2][0]);
        }
        return 0;
    }
    return 0;
}

std::wstring MathManager::CalculateSystemResult(const MathObject& obj)
{
    MathEvaluator eval;
    std::vector<std::wstring> equations;

    // Parse equations from the parts - for system of equations, we use part1 and part2 as separate equations
    // and part3 could be a third equation if present
    if (!obj.SlotText(1).empty()) equations.push_back(obj.SlotText(1));
    if (!obj.SlotText(2).empty()) equations.push_back(obj.SlotText(2));
    if (!obj.SlotText(3).empty()) equations.push_back(obj.SlotText(3));

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
