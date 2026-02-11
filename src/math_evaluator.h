#pragma once

#include <string>
#include <vector>
#include <map>

// Rational number class for exact arithmetic
class Rational {
public:
    long long num;   // numerator
    long long den;   // denominator

    Rational(long long n = 0, long long d = 1) : num(n), den(d) {
        if (den < 0) { num = -num; den = -den; }  // keep denominator positive
        normalize();
    }

    void normalize() {
        if (num == 0) { den = 1; return; }
        long long g = gcd(llabs(num), llabs(den));
        num /= g;
        den /= g;
    }

    static long long gcd(long long a, long long b) {
        while (b != 0) {
            long long t = b;
            b = a % b;
            a = t;
        }
        return a;
    }

    Rational operator+(const Rational& other) const {
        return Rational(num * other.den + other.num * den, den * other.den);
    }

    Rational operator-(const Rational& other) const {
        return Rational(num * other.den - other.num * den, den * other.den);
    }

    Rational operator*(const Rational& other) const {
        return Rational(num * other.num, den * other.den);
    }

    Rational operator/(const Rational& other) const {
        return Rational(num * other.den, den * other.num);
    }

    double toDouble() const {
        return (double)num / den;
    }

    std::wstring toString() const {
        if (den == 1) return std::to_wstring(num);
        return std::to_wstring(num) + L"/" + std::to_wstring(den);
    }
};

class MathEvaluator
{
public:
    // Double-based evaluation methods
    double Eval(const std::wstring& expr, const std::wstring& varName = L"", double varValue = 0);
    std::map<std::wstring, double> SolveSystemOfEquations(const std::vector<std::wstring>& equations);

    // Rational-based evaluation methods
    Rational EvalRational(const std::wstring& expr, const std::wstring& varName = L"", const Rational& varValue = Rational(0));
    std::map<std::wstring, Rational> SolveSystemOfEquationsRational(const std::vector<std::wstring>& equations);

private:
    std::wstring expr;
    size_t pos = 0;
    std::wstring varName;
    
    // Double-based parsing members
    double varValue_d = 0;
    
    // Rational-based parsing members  
    Rational varValue_r;

    // Double-based parsing methods
    double ParseExpression();
    double ParseTerm();
    double ParseFactor();
    double ParsePower();
    
    // Rational-based parsing methods
    Rational ParseExpressionRational();
    Rational ParseTermRational();
    Rational ParseFactorRational();
    Rational ParsePowerRational();
    
    void SkipSpace();
};

bool ParseLowerLimit(const std::wstring& s, std::wstring& var, double& val);
