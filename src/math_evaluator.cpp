#include "math_evaluator.h"
#include <cwctype>
#include <cmath>
#include <cstdlib>
#include <vector>
#include <map>
#include <sstream>
#include <stdexcept>
#include <algorithm>
#include <cctype>

namespace {
    bool IsFactorStart(wchar_t ch)
    {
        return iswdigit(ch) || ch == L'.' || iswalpha(ch) || ch == L'(' || ch == L'{';
    }

    Rational DoubleToRational(double value)
    {
        return Rational((long long)llround(value * 1000000.0), 1000000);
    }

    bool TryApplyUnaryFunction(const std::wstring& name, double arg, double& out)
    {
        if (name == L"sin") { out = sin(arg); return true; }
        if (name == L"cos") { out = cos(arg); return true; }
        if (name == L"tan") { out = tan(arg); return true; }
        if (name == L"asin") { if (arg < -1 || arg > 1) return false; out = asin(arg); return true; }
        if (name == L"acos") { if (arg < -1 || arg > 1) return false; out = acos(arg); return true; }
        if (name == L"atan") { out = atan(arg); return true; }
        if (name == L"sqrt") { if (arg < 0) return false; out = sqrt(arg); return true; }
        if (name == L"abs") { out = fabs(arg); return true; }
        if (name == L"exp") { out = exp(arg); return true; }
        return false;
    }
}

double MathEvaluator::Eval(const std::wstring& e, const std::wstring& vName, double vVal)
{
    expr = e; pos = 0; varName = vName; varValue_d = vVal;
    try { return ParseExpression(); } catch (...) { return 0; }
}

double MathEvaluator::ParseExpression()
{
    double val = ParseTerm();
    while (true)
    {
        SkipSpace();
        if (pos >= expr.size()) break;
        if (expr[pos] == L'+') { pos++; val += ParseTerm(); }
        else if (expr[pos] == L'-') { pos++; val -= ParseTerm(); }
        else break;
    }
    return val;
}

double MathEvaluator::ParseTerm()
{
    double val = ParseFactor();
    while (true)
    {
        SkipSpace();
        if (pos >= expr.size()) break;
        if (expr[pos] == L'*') { pos++; val *= ParseFactor(); }
        else if (expr[pos] == L'/')
        {
            pos++;
            double d = ParseFactor();
            if (d != 0) val /= d;
        }
        else if (IsFactorStart(expr[pos]))
        {
            val *= ParseFactor();
        }
        else break;
    }
    return val;
}

double MathEvaluator::ParseFactor()
{
    double val = ParsePower();
    SkipSpace();
    if (pos < expr.size() && expr[pos] == L'^')
    {
        pos++;
        val = pow(val, ParseFactor());
    }
    return val;
}

double MathEvaluator::ParsePower()
{
    SkipSpace();
    if (pos >= expr.size()) return 0;
    if (expr[pos] == L'(' || expr[pos] == L'{')
    {
        wchar_t close = (expr[pos] == L'(') ? L')' : L'}';
        pos++;
        double val = ParseExpression();
        SkipSpace();
        if (pos < expr.size() && expr[pos] == close) pos++;
        return val;
    }
    if (expr[pos] == L'-') { pos++; return -ParsePower(); }

    if (iswdigit(expr[pos]) || expr[pos] == L'.')
    {
        wchar_t* end;
        double val = wcstod(&expr[pos], &end);
        pos = (size_t)(end - expr.c_str());
        return val;
    }

    if (iswalpha(expr[pos]))
    {
        std::wstring name;
        while (pos < expr.size() && (iswalpha(expr[pos]) || iswdigit(expr[pos])))
        {
            name += expr[pos++];
        }

        if (!varName.empty() && name == varName) return varValue_d;
        if (name == L"pi") return 3.14159265358979;
        if (name == L"e") return 2.718281828459;
        if (name == L"log" || name == L"ln")
        {
            SkipSpace();
            double base = (name == L"ln") ? 2.718281828459 : 10.0;
            // Check for explicit base: log_base(value) or log(base, value)
            if (pos < expr.size() && expr[pos] == L'_')
            {
                pos++;
                // Parse base as number or expression in braces
                SkipSpace();
                if (pos < expr.size() && expr[pos] == L'{')
                {
                    pos++;
                    std::wstring baseExpr;
                    int braceCount = 1;
                    while (pos < expr.size() && braceCount > 0)
                    {
                        if (expr[pos] == L'{') braceCount++;
                        else if (expr[pos] == L'}') braceCount--;
                        if (braceCount > 0) baseExpr += expr[pos];
                        pos++;
                    }
                    MathEvaluator subEval;
                    base = subEval.Eval(baseExpr, varName, varValue_d);
                }
                else
                {
                    // Parse number
                    std::wstring baseStr;
                    while (pos < expr.size() && (iswdigit(expr[pos]) || expr[pos] == L'.'))
                    {
                        baseStr += expr[pos++];
                    }
                    base = _wtof(baseStr.c_str());
                }
                SkipSpace();
            }
            // Parse argument
            SkipSpace();
            if (pos < expr.size() && expr[pos] == L'(')
            {
                pos++;
                double arg = ParseExpression();
                SkipSpace();
                if (pos < expr.size() && expr[pos] == L')') pos++;
                if (arg > 0 && base > 0 && base != 1)
                    return log(arg) / log(base);
            }
            else if (pos < expr.size() && expr[pos] == L'{')
            {
                pos++;
                double arg = ParseExpression();
                SkipSpace();
                if (pos < expr.size() && expr[pos] == L'}') pos++;
                if (arg > 0 && base > 0 && base != 1)
                    return log(arg) / log(base);
            }
            return 0;
        }

        SkipSpace();
        if (pos < expr.size() && (expr[pos] == L'(' || expr[pos] == L'{'))
        {
            wchar_t close = (expr[pos] == L'(') ? L')' : L'}';
            pos++;
            double arg = ParseExpression();
            SkipSpace();
            if (pos < expr.size() && expr[pos] == close) pos++;

            double funcResult = 0;
            if (TryApplyUnaryFunction(name, arg, funcResult))
                return funcResult;
            return 0;
        }
    }
    return 0;
}

Rational MathEvaluator::EvalRational(const std::wstring& e, const std::wstring& vName, const Rational& vVal)
{
    expr = e; pos = 0; varName = vName; varValue_r = vVal;
    try { return ParseExpressionRational(); } catch (...) { return Rational(0); }
}

Rational MathEvaluator::ParseExpressionRational()
{
    Rational val = ParseTermRational();
    while (true)
    {
        SkipSpace();
        if (pos >= expr.size()) break;
        if (expr[pos] == L'+') { pos++; val = val + ParseTermRational(); }
        else if (expr[pos] == L'-') { pos++; val = val - ParseTermRational(); }
        else break;
    }
    return val;
}

Rational MathEvaluator::ParseTermRational()
{
    Rational val = ParseFactorRational();
    while (true)
    {
        SkipSpace();
        if (pos >= expr.size()) break;
        if (expr[pos] == L'*') { pos++; val = val * ParseFactorRational(); }
        else if (expr[pos] == L'/')
        {
            pos++;
            Rational d = ParseFactorRational();
            if (d.num != 0) val = val / d;
        }
        else if (IsFactorStart(expr[pos]))
        {
            val = val * ParseFactorRational();
        }
        else break;
    }
    return val;
}

Rational MathEvaluator::ParseFactorRational()
{
    Rational val = ParsePowerRational();
    SkipSpace();
    if (pos < expr.size() && expr[pos] == L'^')
    {
        pos++;
        // For rational arithmetic, we'll only handle integer exponents
        Rational exp = ParseFactorRational();
        if (exp.den == 1) {  // integer exponent
            long long n = exp.num;
            if (n == 0) return Rational(1);
            else if (n > 0) {
                Rational result(1);
                for (long long i = 0; i < n; i++) {
                    result = result * val;
                }
                return result;
            } else {  // negative exponent
                Rational base_inv = Rational(val.den, val.num);  // reciprocal
                Rational result(1);
                for (long long i = 0; i < -n; i++) {
                    result = result * base_inv;
                }
                return result;
            }
        }
        // If not integer exponent, return the double result as a rational approximation
        return Rational((long long)(pow(val.toDouble(), exp.toDouble()) * 1000000), 1000000);
    }
    return val;
}

Rational MathEvaluator::ParsePowerRational()
{
    SkipSpace();
    if (pos >= expr.size()) return Rational(0);
    if (expr[pos] == L'(' || expr[pos] == L'{')
    {
        wchar_t close = (expr[pos] == L'(') ? L')' : L'}';
        pos++;
        Rational val = ParseExpressionRational();
        SkipSpace();
        if (pos < expr.size() && expr[pos] == close) pos++;
        return val;
    }
    if (expr[pos] == L'-') { pos++; return Rational(0) - ParsePowerRational(); }

    if (iswdigit(expr[pos]) || expr[pos] == L'.')
    {
        wchar_t* end;
        double val = wcstod(&expr[pos], &end);
        pos = (size_t)(end - expr.c_str());
        return DoubleToRational(val);
    }

    if (iswalpha(expr[pos]))
    {
        std::wstring name;
        while (pos < expr.size() && (iswalpha(expr[pos]) || iswdigit(expr[pos])))
        {
            name += expr[pos++];
        }
        if (!varName.empty() && name == varName) return varValue_r;
        if (name == L"pi") return Rational(3141592, 1000000);  // Approximation of pi
        if (name == L"e") return Rational(2718281, 1000000);  // Approximation of e

        if (name == L"log" || name == L"ln")
        {
            SkipSpace();
            double base = (name == L"ln") ? 2.718281828459 : 10.0;

            if (pos < expr.size() && expr[pos] == L'_')
            {
                pos++;
                SkipSpace();
                if (pos < expr.size() && expr[pos] == L'{')
                {
                    pos++;
                    std::wstring baseExpr;
                    int braceCount = 1;
                    while (pos < expr.size() && braceCount > 0)
                    {
                        if (expr[pos] == L'{') braceCount++;
                        else if (expr[pos] == L'}') braceCount--;
                        if (braceCount > 0) baseExpr += expr[pos];
                        pos++;
                    }
                    MathEvaluator subEval;
                    base = subEval.EvalRational(baseExpr, varName, varValue_r).toDouble();
                }
                else
                {
                    std::wstring baseStr;
                    while (pos < expr.size() && (iswdigit(expr[pos]) || expr[pos] == L'.'))
                    {
                        baseStr += expr[pos++];
                    }
                    base = _wtof(baseStr.c_str());
                }
                SkipSpace();
            }

            SkipSpace();
            if (pos < expr.size() && (expr[pos] == L'(' || expr[pos] == L'{'))
            {
                wchar_t close = (expr[pos] == L'(') ? L')' : L'}';
                pos++;
                double arg = ParseExpressionRational().toDouble();
                SkipSpace();
                if (pos < expr.size() && expr[pos] == close) pos++;
                if (arg > 0 && base > 0 && base != 1)
                    return DoubleToRational(log(arg) / log(base));
            }
            return Rational(0);
        }

        SkipSpace();
        if (pos < expr.size() && (expr[pos] == L'(' || expr[pos] == L'{'))
        {
            wchar_t close = (expr[pos] == L'(') ? L')' : L'}';
            pos++;
            double arg = ParseExpressionRational().toDouble();
            SkipSpace();
            if (pos < expr.size() && expr[pos] == close) pos++;

            double funcResult = 0;
            if (TryApplyUnaryFunction(name, arg, funcResult))
                return DoubleToRational(funcResult);
            return Rational(0);
        }
    }
    return Rational(0);
}

void MathEvaluator::SkipSpace() { while (pos < expr.size() && iswspace(expr[pos])) pos++; }

bool ParseLowerLimit(const std::wstring& s, std::wstring& var, double& val)
{
    size_t eq = s.find(L'=');
    if (eq == std::wstring::npos) { var = L"i"; val = _wtof(s.c_str()); return true; }
    var = s.substr(0, eq);
    size_t first = var.find_first_not_of(L" ");
    size_t last = var.find_last_not_of(L" ");
    if (first != std::wstring::npos && last != std::wstring::npos)
        var = var.substr(first, last - first + 1);
    val = _wtof(s.substr(eq + 1).c_str());
    return true;
}

// Helper class for linear equation parsing with rationals
struct LinearEquationRational {
    Rational x_coeff = Rational(0);
    Rational y_coeff = Rational(0);
    Rational z_coeff = Rational(0);
    Rational constant = Rational(0);
};

// Parse a linear equation into rational coefficients
LinearEquationRational ParseLinearEquationRational(const std::wstring& equation) {
    LinearEquationRational result;
    MathEvaluator eval;

    size_t eq_pos = equation.find(L'=');
    if (eq_pos == std::wstring::npos) {
        throw std::runtime_error("Equation must contain '='");
    }

    std::wstring left_side = equation.substr(0, eq_pos);
    std::wstring right_side = equation.substr(eq_pos + 1);

    // Evaluate right side as constant using rational arithmetic
    Rational right_value = eval.EvalRational(right_side);

    // Parse left side to get coefficients
    std::wstring term;
    Rational sign(1);
    size_t pos = 0;

    while (pos < left_side.size()) {
        // Skip spaces
        while (pos < left_side.size() && iswspace(left_side[pos])) pos++;
        if (pos >= left_side.size()) break;

        // Check for sign
        if (left_side[pos] == L'+') {
            sign = Rational(1);
            pos++;
        } else if (left_side[pos] == L'-') {
            sign = Rational(-1);
            pos++;
        }

        // Skip spaces after sign
        while (pos < left_side.size() && iswspace(left_side[pos])) pos++;
        if (pos >= left_side.size()) break;

        // Parse term
        term.clear();
        while (pos < left_side.size() && !iswspace(left_side[pos]) &&
               left_side[pos] != L'+' && left_side[pos] != L'-') {
            term += left_side[pos++];
        }

        if (term.empty()) continue;

        // Check if term contains a variable
        bool has_variable = false;
        for (wchar_t c : term) {
            if (iswalpha(c) && c != L'.') {
                has_variable = true;
                break;
            }
        }

        if (has_variable) {
            // Parse coefficient and variable
            std::wstring coeff_str, var_name;
            for (wchar_t c : term) {
                if (iswalpha(c) && c != L'.') {
                    var_name += c;
                } else {
                    coeff_str += c;
                }
            }

            Rational coefficient(1);
            if (!coeff_str.empty()) {
                if (coeff_str == L"-") coefficient = Rational(-1);
                else if (coeff_str == L"+") coefficient = Rational(1);
                else coefficient = eval.EvalRational(coeff_str);
            }

            coefficient = coefficient * sign;

            if (var_name == L"x") result.x_coeff = result.x_coeff + coefficient;
            else if (var_name == L"y") result.y_coeff = result.y_coeff + coefficient;
            else if (var_name == L"z") result.z_coeff = result.z_coeff + coefficient;
        } else {
            // Constant term on left side
            Rational constant = eval.EvalRational(term) * sign;
            result.constant = result.constant - constant; // Move to right side
        }
    }

    result.constant = result.constant + right_value;
    return result;
}

// Solve 2x2 linear system using rational arithmetic
std::map<std::wstring, Rational> Solve2x2SystemRational(const Rational& a1, const Rational& b1, const Rational& c1,
                                                        const Rational& a2, const Rational& b2, const Rational& c2) {
    std::map<std::wstring, Rational> result;

    // Calculate determinant: a1*b2 - a2*b1
    Rational det = a1 * b2 - a2 * b1;
    
    if (det.num == 0) {  // Determinant is zero
        // Check if system is inconsistent or has infinite solutions
        Rational check1 = a1 * c2 - a2 * c1;
        Rational check2 = b1 * c2 - b2 * c1;
        
        if (check1.num == 0 && check2.num == 0) {
            result[L"x"] = Rational(0);
            result[L"y"] = Rational(0);
            result[L"status"] = Rational(-1); // Infinite solutions
        } else {
            result[L"x"] = Rational(0);
            result[L"y"] = Rational(0);
            result[L"status"] = Rational(-2); // No solution
        }
        return result;
    }

    // Calculate solutions: x = (c1*b2 - c2*b1) / det, y = (a1*c2 - a2*c1) / det
    result[L"x"] = (c1 * b2 - c2 * b1) / det;
    result[L"y"] = (a1 * c2 - a2 * c1) / det;
    result[L"status"] = Rational(0); // Success
    return result;
}

// Solve 3x3 linear system using rational arithmetic
std::map<std::wstring, Rational> Solve3x3SystemRational(const Rational& a1, const Rational& b1, const Rational& c1, const Rational& d1,
                                                        const Rational& a2, const Rational& b2, const Rational& c2, const Rational& d2,
                                                        const Rational& a3, const Rational& b3, const Rational& c3, const Rational& d3) {
    std::map<std::wstring, Rational> result;

    // Calculate determinant
    Rational det = a1 * (b2 * c3 - b3 * c2) - b1 * (a2 * c3 - a3 * c2) + c1 * (a2 * b3 - a3 * b2);

    if (det.num == 0) {
        result[L"x"] = Rational(0);
        result[L"y"] = Rational(0);
        result[L"z"] = Rational(0);
        result[L"status"] = Rational(-2); // No unique solution
        return result;
    }

    // Calculate determinants for x, y, z
    Rational det_x = d1 * (b2 * c3 - b3 * c2) - b1 * (d2 * c3 - d3 * c2) + c1 * (d2 * b3 - d3 * b2);
    Rational det_y = a1 * (d2 * c3 - d3 * c2) - d1 * (a2 * c3 - a3 * c2) + c1 * (a2 * d3 - a3 * d2);
    Rational det_z = a1 * (b2 * d3 - b3 * d2) - b1 * (a2 * d3 - a3 * d2) + d1 * (a2 * b3 - a3 * b2);

    result[L"x"] = det_x / det;
    result[L"y"] = det_y / det;
    result[L"z"] = det_z / det;
    result[L"status"] = Rational(0); // Success
    return result;
}

std::map<std::wstring, Rational> MathEvaluator::SolveSystemOfEquationsRational(const std::vector<std::wstring>& equations) {
    if (equations.empty()) {
        std::map<std::wstring, Rational> result;
        result[L"status"] = Rational(-3); // No equations
        return result;
    }

    // Parse all equations
    std::vector<LinearEquationRational> parsed_equations;
    for (const auto& eq : equations) {
        try {
            parsed_equations.push_back(ParseLinearEquationRational(eq));
        } catch (...) {
            std::map<std::wstring, Rational> result;
            result[L"status"] = Rational(-4); // Parse error
            return result;
        }
    }

    size_t num_equations = parsed_equations.size();

    if (num_equations == 1) {
        // Single equation - solve for one variable if possible
        const auto& eq = parsed_equations[0];
        std::map<std::wstring, Rational> result;

        if (eq.x_coeff.num != 0 && eq.y_coeff.num == 0 && eq.z_coeff.num == 0) {
            result[L"x"] = eq.constant / eq.x_coeff;
            result[L"y"] = Rational(0);
            result[L"z"] = Rational(0);
            result[L"status"] = Rational(0);
        } else if (eq.y_coeff.num != 0 && eq.x_coeff.num == 0 && eq.z_coeff.num == 0) {
            result[L"x"] = Rational(0);
            result[L"y"] = eq.constant / eq.y_coeff;
            result[L"z"] = Rational(0);
            result[L"status"] = Rational(0);
        } else if (eq.z_coeff.num != 0 && eq.x_coeff.num == 0 && eq.y_coeff.num == 0) {
            result[L"x"] = Rational(0);
            result[L"y"] = Rational(0);
            result[L"z"] = eq.constant / eq.z_coeff;
            result[L"status"] = Rational(0);
        } else {
            result[L"status"] = Rational(-5); // Underdetermined
        }
        return result;
    }

    if (num_equations == 2) {
        // 2 equations - solve 2x2 system
        const auto& eq1 = parsed_equations[0];
        const auto& eq2 = parsed_equations[1];

        // Check which variables are present
        bool has_x = eq1.x_coeff.num != 0 || eq2.x_coeff.num != 0;
        bool has_y = eq1.y_coeff.num != 0 || eq2.y_coeff.num != 0;
        bool has_z = eq1.z_coeff.num != 0 || eq2.z_coeff.num != 0;

        if (has_z && !has_x && !has_y) {
            // Solve for z only
            return Solve2x2SystemRational(eq1.z_coeff, Rational(0), eq1.constant,
                                         eq2.z_coeff, Rational(0), eq2.constant);
        } else if (has_z && has_x && !has_y) {
            // Solve for x and z
            auto result = Solve2x2SystemRational(eq1.x_coeff, eq1.z_coeff, eq1.constant,
                                                eq2.x_coeff, eq2.z_coeff, eq2.constant);
            result[L"y"] = Rational(0);
            return result;
        } else if (has_z && has_y && !has_x) {
            // Solve for y and z
            auto result = Solve2x2SystemRational(eq1.y_coeff, eq1.z_coeff, eq1.constant,
                                                eq2.y_coeff, eq2.z_coeff, eq2.constant);
            result[L"x"] = Rational(0);
            return result;
        } else {
            // Solve for x and y (ignore z)
            return Solve2x2SystemRational(eq1.x_coeff, eq1.y_coeff, eq1.constant,
                                         eq2.x_coeff, eq2.y_coeff, eq2.constant);
        }
    }

    if (num_equations == 3) {
        // 3 equations - solve 3x3 system
        const auto& eq1 = parsed_equations[0];
        const auto& eq2 = parsed_equations[1];
        const auto& eq3 = parsed_equations[2];

        return Solve3x3SystemRational(eq1.x_coeff, eq1.y_coeff, eq1.z_coeff, eq1.constant,
                                     eq2.x_coeff, eq2.y_coeff, eq2.z_coeff, eq2.constant,
                                     eq3.x_coeff, eq3.y_coeff, eq3.z_coeff, eq3.constant);
    }

    // More than 3 equations - not supported
    std::map<std::wstring, Rational> result;
    result[L"status"] = Rational(-6);
    return result;
}

// Original double-based functions still available for compatibility
// Helper class for linear equation parsing
struct LinearEquation {
    double x_coeff = 0;
    double y_coeff = 0;
    double z_coeff = 0;
    double constant = 0;
};

// Parse a linear equation into coefficients
LinearEquation ParseLinearEquation(const std::wstring& equation) {
    LinearEquation result;
    MathEvaluator eval;

    size_t eq_pos = equation.find(L'=');
    if (eq_pos == std::wstring::npos) {
        throw std::runtime_error("Equation must contain '='");
    }

    std::wstring left_side = equation.substr(0, eq_pos);
    std::wstring right_side = equation.substr(eq_pos + 1);

    // Evaluate right side as constant
    double right_value = eval.Eval(right_side);

    // Parse left side to get coefficients
    std::wstring term;
    double sign = 1.0;
    size_t pos = 0;

    while (pos < left_side.size()) {
        // Skip spaces
        while (pos < left_side.size() && iswspace(left_side[pos])) pos++;
        if (pos >= left_side.size()) break;

        // Check for sign
        if (left_side[pos] == L'+') {
            sign = 1.0;
            pos++;
        } else if (left_side[pos] == L'-') {
            sign = -1.0;
            pos++;
        }

        // Skip spaces after sign
        while (pos < left_side.size() && iswspace(left_side[pos])) pos++;
        if (pos >= left_side.size()) break;

        // Parse term
        term.clear();
        while (pos < left_side.size() && !iswspace(left_side[pos]) &&
               left_side[pos] != L'+' && left_side[pos] != L'-') {
            term += left_side[pos++];
        }

        if (term.empty()) continue;

        // Check if term contains a variable
        bool has_variable = false;
        for (wchar_t c : term) {
            if (iswalpha(c) && c != L'.') {
                has_variable = true;
                break;
            }
        }

        if (has_variable) {
            // Parse coefficient and variable
            std::wstring coeff_str, var_name;
            for (wchar_t c : term) {
                if (iswalpha(c) && c != L'.') {
                    var_name += c;
                } else {
                    coeff_str += c;
                }
            }

            double coefficient = 1.0;
            if (!coeff_str.empty()) {
                if (coeff_str == L"-") coefficient = -1.0;
                else if (coeff_str == L"+") coefficient = 1.0;
                else coefficient = eval.Eval(coeff_str);
            }

            coefficient *= sign;

            if (var_name == L"x") result.x_coeff += coefficient;
            else if (var_name == L"y") result.y_coeff += coefficient;
            else if (var_name == L"z") result.z_coeff += coefficient;
        } else {
            // Constant term on left side
            double constant = eval.Eval(term) * sign;
            result.constant -= constant; // Move to right side
        }
    }

    result.constant += right_value;
    return result;
}

// Solve 2x2 linear system
std::map<std::wstring, double> Solve2x2System(double a1, double b1, double c1,
                                             double a2, double b2, double c2) {
    std::map<std::wstring, double> result;

    double det = a1 * b2 - a2 * b1;
    if (fabs(det) < 1e-10) {
        // Check if system is inconsistent or has infinite solutions
        if (fabs(a1 * c2 - a2 * c1) < 1e-10 && fabs(b1 * c2 - b2 * c1) < 1e-10) {
            result[L"x"] = 0;
            result[L"y"] = 0;
            result[L"status"] = -1; // Infinite solutions
        } else {
            result[L"x"] = 0;
            result[L"y"] = 0;
            result[L"status"] = -2; // No solution
        }
        return result;
    }

    result[L"x"] = (c1 * b2 - c2 * b1) / det;
    result[L"y"] = (a1 * c2 - a2 * c1) / det;
    result[L"status"] = 0; // Success
    return result;
}

// Solve 3x3 linear system
std::map<std::wstring, double> Solve3x3System(double a1, double b1, double c1, double d1,
                                             double a2, double b2, double c2, double d2,
                                             double a3, double b3, double c3, double d3) {
    std::map<std::wstring, double> result;

    double det = a1 * (b2 * c3 - b3 * c2)
               - b1 * (a2 * c3 - a3 * c2)
               + c1 * (a2 * b3 - a3 * b2);

    if (fabs(det) < 1e-10) {
        // Try to solve as 2x2 system if possible
        // For simplicity, return zeros with error status
        result[L"x"] = 0;
        result[L"y"] = 0;
        result[L"z"] = 0;
        result[L"status"] = -2; // No unique solution
        return result;
    }

    double det_x = d1 * (b2 * c3 - b3 * c2)
                 - b1 * (d2 * c3 - d3 * c2)
                 + c1 * (d2 * b3 - d3 * b2);

    double det_y = a1 * (d2 * c3 - d3 * c2)
                 - d1 * (a2 * c3 - a3 * c2)
                 + c1 * (a2 * d3 - a3 * d2);

    double det_z = a1 * (b2 * d3 - b3 * d2)
                 - b1 * (a2 * d3 - a3 * d2)
                 + d1 * (a2 * b3 - a3 * b2);

    result[L"x"] = det_x / det;
    result[L"y"] = det_y / det;
    result[L"z"] = det_z / det;
    result[L"status"] = 0; // Success
    return result;
}

std::map<std::wstring, double> MathEvaluator::SolveSystemOfEquations(const std::vector<std::wstring>& equations) {
    if (equations.empty()) {
        return {{L"status", -3}}; // No equations
    }

    // Parse all equations
    std::vector<LinearEquation> parsed_equations;
    for (const auto& eq : equations) {
        try {
            parsed_equations.push_back(ParseLinearEquation(eq));
        } catch (...) {
            return {{L"status", -4}}; // Parse error
        }
    }

    size_t num_equations = parsed_equations.size();

    if (num_equations == 1) {
        // Single equation - solve for one variable if possible
        const auto& eq = parsed_equations[0];
        std::map<std::wstring, double> result;

        if (fabs(eq.x_coeff) > 1e-10 && fabs(eq.y_coeff) < 1e-10 && fabs(eq.z_coeff) < 1e-10) {
            result[L"x"] = eq.constant / eq.x_coeff;
            result[L"y"] = 0;
            result[L"z"] = 0;
            result[L"status"] = 0;
        } else if (fabs(eq.y_coeff) > 1e-10 && fabs(eq.x_coeff) < 1e-10 && fabs(eq.z_coeff) < 1e-10) {
            result[L"x"] = 0;
            result[L"y"] = eq.constant / eq.y_coeff;
            result[L"z"] = 0;
            result[L"status"] = 0;
        } else if (fabs(eq.z_coeff) > 1e-10 && fabs(eq.x_coeff) < 1e-10 && fabs(eq.y_coeff) < 1e-10) {
            result[L"x"] = 0;
            result[L"y"] = 0;
            result[L"z"] = eq.constant / eq.z_coeff;
            result[L"status"] = 0;
        } else {
            result[L"status"] = -5; // Underdetermined
        }
        return result;
    }

    if (num_equations == 2) {
        // 2 equations - solve 2x2 system
        const auto& eq1 = parsed_equations[0];
        const auto& eq2 = parsed_equations[1];

        // Check which variables are present
        bool has_x = fabs(eq1.x_coeff) > 1e-10 || fabs(eq2.x_coeff) > 1e-10;
        bool has_y = fabs(eq1.y_coeff) > 1e-10 || fabs(eq2.y_coeff) > 1e-10;
        bool has_z = fabs(eq1.z_coeff) > 1e-10 || fabs(eq2.z_coeff) > 1e-10;

        if (has_z && !has_x && !has_y) {
            // Solve for z only
            return Solve2x2System(eq1.z_coeff, 0, eq1.constant,
                                 eq2.z_coeff, 0, eq2.constant);
        } else if (has_z && has_x && !has_y) {
            // Solve for x and z
            auto result = Solve2x2System(eq1.x_coeff, eq1.z_coeff, eq1.constant,
                                        eq2.x_coeff, eq2.z_coeff, eq2.constant);
            result[L"y"] = 0;
            return result;
        } else if (has_z && has_y && !has_x) {
            // Solve for y and z
            auto result = Solve2x2System(eq1.y_coeff, eq1.z_coeff, eq1.constant,
                                        eq2.y_coeff, eq2.z_coeff, eq2.constant);
            result[L"x"] = 0;
            return result;
        } else {
            // Solve for x and y (ignore z)
            return Solve2x2System(eq1.x_coeff, eq1.y_coeff, eq1.constant,
                                 eq2.x_coeff, eq2.y_coeff, eq2.constant);
        }
    }

    if (num_equations == 3) {
        // 3 equations - solve 3x3 system
        const auto& eq1 = parsed_equations[0];
        const auto& eq2 = parsed_equations[1];
        const auto& eq3 = parsed_equations[2];

        return Solve3x3System(eq1.x_coeff, eq1.y_coeff, eq1.z_coeff, eq1.constant,
                             eq2.x_coeff, eq2.y_coeff, eq2.z_coeff, eq2.constant,
                             eq3.x_coeff, eq3.y_coeff, eq3.z_coeff, eq3.constant);
    }

    // More than 3 equations - not supported
    return {{L"status", -6}};
}
