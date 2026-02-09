#include "math_evaluator.h"
#include <cwctype>
#include <cmath>
#include <cstdlib>

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
