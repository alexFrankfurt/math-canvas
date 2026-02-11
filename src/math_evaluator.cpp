#include "math_evaluator.h"
#include <cwctype>
#include <cmath>
#include <cstdlib>
#include <vector>
#include <map>
#include <sstream>
#include <stdexcept>

double MathEvaluator::Eval(const std::wstring& e, const std::wstring& vName, double vVal)
{
    expr = e; pos = 0; varName = vName; varValue = vVal;
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
        if (!varName.empty() && name == varName) return varValue;
        if (name == L"pi") return 3.14159265358979;
        if (name == L"e") return 2.718281828459;
    }
    return 0;
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
