#pragma once

#include <string>
#include <vector>
#include <map>

class MathEvaluator
{
public:
    double Eval(const std::wstring& expr, const std::wstring& varName = L"", double varValue = 0);
    std::map<std::wstring, double> SolveSystemOfEquations(const std::vector<std::wstring>& equations);

private:
    std::wstring expr;
    size_t pos = 0;
    double varValue = 0;
    std::wstring varName;

    double ParseExpression();
    double ParseTerm();
    double ParseFactor();
    double ParsePower();
    void SkipSpace();
};

bool ParseLowerLimit(const std::wstring& s, std::wstring& var, double& val);
