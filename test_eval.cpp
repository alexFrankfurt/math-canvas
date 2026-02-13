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

    run(CheckNearVar(eval, L"2x+1", L"x", 4.0, 9.0, L"2x+1 at x=4"));
    run(CheckNearVar(eval, L"x(x+1)", L"x", 3.0, 12.0, L"x(x+1) at x=3"));
    run(CheckNearVar(eval, L"3(x+2)", L"x", 5.0, 21.0, L"3(x+2) at x=5"));

    run(CheckZero(eval, L"unknown(5)", L"unknown function -> 0"));
    run(CheckZero(eval, L")", L"bad token -> 0"));
    run(CheckZero(eval, L"log_0(10)", L"log base 0 -> 0"));
    run(CheckZero(eval, L"ln(0)", L"ln(0) -> 0"));

    std::wcout << L"\n=== Summary ===" << std::endl;
    std::wcout << L"Passed: " << passed << std::endl;
    std::wcout << L"Failed: " << failed << std::endl;

    return (failed == 0) ? 0 : 1;
}