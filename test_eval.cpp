#include <cmath>
#include <iostream>
#include <string>
#include "src/math_evaluator.h"

namespace {
    constexpr double kEps = 1e-6;

    bool NearlyEqual(double actual, double expected, double eps = kEps)
    {
        return std::fabs(actual - expected) <= eps;
    }

    bool CheckNear(MathEvaluator& eval, const std::wstring& expr, double expected, const std::wstring& label)
    {
        double actual = eval.Eval(expr);
        bool ok = NearlyEqual(actual, expected);
        std::wcout << (ok ? L"[PASS] " : L"[FAIL] ")
                   << label << L" | expr=" << expr
                   << L" | expected=" << expected
                   << L" | actual=" << actual << std::endl;
        return ok;
    }

    bool CheckNearVar(MathEvaluator& eval,
                      const std::wstring& expr,
                      const std::wstring& varName,
                      double varValue,
                      double expected,
                      const std::wstring& label)
    {
        double actual = eval.Eval(expr, varName, varValue);
        bool ok = NearlyEqual(actual, expected);
        std::wcout << (ok ? L"[PASS] " : L"[FAIL] ")
                   << label << L" | expr=" << expr
                   << L" | " << varName << L"=" << varValue
                   << L" | expected=" << expected
                   << L" | actual=" << actual << std::endl;
        return ok;
    }

    bool CheckZero(MathEvaluator& eval, const std::wstring& expr, const std::wstring& label)
    {
        double actual = eval.Eval(expr);
        bool ok = NearlyEqual(actual, 0.0);
        std::wcout << (ok ? L"[PASS] " : L"[FAIL] ")
                   << label << L" | expr=" << expr
                   << L" | expected=0"
                   << L" | actual=" << actual << std::endl;
        return ok;
    }

    bool CheckValue(MathEvaluator& eval,
                    const std::wstring& expr,
                    double expectedDisplayValue,
                    const std::wstring& expectedDisplayUnit,
                    const std::wstring& label)
    {
        const MathValue actual = eval.EvalValue(expr);
        const double scale = std::fabs(actual.displayScale) < 1e-12 ? 1.0 : actual.displayScale;
        const double displayedValue = actual.baseValue / scale;
        const bool ok = !actual.IsError() && NearlyEqual(displayedValue, expectedDisplayValue) && actual.displayUnit == expectedDisplayUnit;

        std::wcout << (ok ? L"[PASS] " : L"[FAIL] ")
                   << label << L" | expr=" << expr
                   << L" | expected=" << expectedDisplayValue;
        if (!expectedDisplayUnit.empty())
            std::wcout << L" " << expectedDisplayUnit;

        if (actual.IsError())
        {
            std::wcout << L" | actual error=" << actual.errorText << std::endl;
        }
        else
        {
            std::wcout << L" | actual=" << displayedValue;
            if (!actual.displayUnit.empty())
                std::wcout << L" " << actual.displayUnit;
            std::wcout << std::endl;
        }

        return ok;
    }

    bool CheckValueError(MathEvaluator& eval,
                         const std::wstring& expr,
                         const std::wstring& expectedError,
                         const std::wstring& label)
    {
        const MathValue actual = eval.EvalValue(expr);
        const bool ok = actual.IsError() && actual.errorText == expectedError;
        std::wcout << (ok ? L"[PASS] " : L"[FAIL] ")
                   << label << L" | expr=" << expr
                   << L" | expected error=" << expectedError
                   << L" | actual=" << (actual.IsError() ? actual.errorText : L"<value>") << std::endl;
        return ok;
    }
}

int main()
{
    MathEvaluator eval;
    int passed = 0;
    int failed = 0;

    auto run = [&](bool ok) {
        if (ok) ++passed;
        else ++failed;
    };

    std::wcout << L"=== Math Evaluator Tests ===" << std::endl;

    run(CheckNear(eval, L"sin(pi/2)", 1.0, L"sin(pi/2)"));
    run(CheckNear(eval, L"cos(0)", 1.0, L"cos(0)"));
    run(CheckNear(eval, L"tan(0)", 0.0, L"tan(0)"));
    run(CheckNear(eval, L"exp(1)", 2.718281828459, L"exp(1)"));
    run(CheckNear(eval, L"sqrt(81)", 9.0, L"sqrt(81)"));
    run(CheckNear(eval, L"abs(-12.5)", 12.5, L"abs(-12.5)"));

    run(CheckNear(eval, L"ln(e)", 1.0, L"ln(e)"));
    run(CheckNear(eval, L"log(100)", 2.0, L"log base-10"));
    run(CheckNear(eval, L"log_2(8)", 3.0, L"log_2(8)"));
    run(CheckNear(eval, L"log_{2}(8)", 3.0, L"log_{2}(8)"));
    run(CheckZero(eval, L"log_1(10)", L"invalid log base -> 0"));
    run(CheckZero(eval, L"log(-10)", L"log invalid arg -> 0"));
    run(CheckZero(eval, L"ln(-1)", L"ln invalid arg -> 0"));

    run(CheckNear(eval, L"2(3+4)", 14.0, L"number-parenthesis implicit multiplication"));
    run(CheckNear(eval, L"(1+2)(3+4)", 21.0, L"parenthesis-parenthesis implicit multiplication"));
    run(CheckNear(eval, L"3pi", 3.0 * 3.14159265358979, L"number-constant implicit multiplication"));
    run(CheckNear(eval, L"sqrt(((16)/(4)))", 2.0, L"nested structured fraction flattening"));
    run(CheckNear(eval, L"abs(-5+((2)^(4)))", 11.0, L"nested structured power flattening"));
    run(CheckNear(eval, L"log_{2}(((3)^(2)))", std::log(9.0) / std::log(2.0), L"nested structured logarithm flattening"));

    run(CheckNearVar(eval, L"2x+1", L"x", 4.0, 9.0, L"2x+1 at x=4"));
    run(CheckNearVar(eval, L"x(x+1)", L"x", 3.0, 12.0, L"x(x+1) at x=3"));
    run(CheckNearVar(eval, L"3(x+2)", L"x", 5.0, 21.0, L"3(x+2) at x=5"));

    run(CheckValue(eval, L"3m + 40cm", 3.4, L"m", L"compatible unit addition converts to left display unit"));
    run(CheckValue(eval, L"30cm + 0.7m", 100.0, L"cm", L"compatible unit addition preserves left display unit"));
    run(CheckValue(eval, L"5kg * 2m / s^2", 10.0, L"N", L"derived unit composition normalizes to Newtons"));
    run(CheckValue(eval, L"sqrt(9m^2)", 3.0, L"m", L"square root reduces even unit exponents"));
    run(CheckValueError(eval, L"3m + 2s", L"incompatible units", L"incompatible unit addition surfaces explicit error"));
    run(CheckValueError(eval, L"sqrt(9m)", L"invalid unit exponent", L"invalid unit exponent is rejected"));
    run(CheckValueError(eval, L"log(10m)", L"log requires abstract number", L"logarithm rejects dimensional quantities"));

    run(CheckZero(eval, L"unknown(5)", L"unknown function -> 0"));
    run(CheckZero(eval, L")", L"bad token -> 0"));
    run(CheckZero(eval, L"log_0(10)", L"log base 0 -> 0"));
    run(CheckZero(eval, L"ln(0)", L"ln(0) -> 0"));

    std::wcout << L"\n=== Summary ===" << std::endl;
    std::wcout << L"Passed: " << passed << std::endl;
    std::wcout << L"Failed: " << failed << std::endl;

    return (failed == 0) ? 0 : 1;
}