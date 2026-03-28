#pragma once

#include <string>
#include <vector>
#include <map>

struct UnitDimension {
    int length = 0;
    int mass = 0;
    int time = 0;
    int current = 0;
    int temperature = 0;
    int amount = 0;
    int luminousIntensity = 0;

    bool IsDimensionless() const {
        return length == 0 && mass == 0 && time == 0 && current == 0 &&
               temperature == 0 && amount == 0 && luminousIntensity == 0;
    }

    bool operator==(const UnitDimension& other) const {
        return length == other.length && mass == other.mass && time == other.time &&
               current == other.current && temperature == other.temperature &&
               amount == other.amount && luminousIntensity == other.luminousIntensity;
    }

    bool operator!=(const UnitDimension& other) const {
        return !(*this == other);
    }
};

struct MathValue {
    double baseValue = 0.0;
    UnitDimension dimension = {};
    double displayScale = 1.0;
    std::wstring displayUnit;
    std::wstring errorText;

    static MathValue Scalar(double value) {
        MathValue result;
        result.baseValue = value;
        return result;
    }

    static MathValue Quantity(double value, const UnitDimension& dim, double scale, const std::wstring& unit) {
        MathValue result;
        result.baseValue = value;
        result.dimension = dim;
        result.displayScale = scale;
        result.displayUnit = unit;
        return result;
    }

    static MathValue Error(const std::wstring& message) {
        MathValue result;
        result.errorText = message;
        return result;
    }

    bool IsError() const {
        return !errorText.empty();
    }

    bool IsDimensionless() const {
        return dimension.IsDimensionless();
    }

    bool HasDisplayUnit() const {
        return !displayUnit.empty();
    }
};

std::wstring BuildCanonicalUnitSymbol(const UnitDimension& dimension);
const std::vector<std::wstring>& GetKnownUnitSymbols();
std::vector<std::wstring> FindMatchingUnitSymbols(const std::wstring& prefix);

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
    MathValue EvalValue(const std::wstring& expr, const std::wstring& varName = L"", const MathValue& varValue = MathValue::Scalar(0.0));
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

    // Quantity-based parsing members
    MathValue varValue_q;
    
    // Rational-based parsing members  
    Rational varValue_r;

    // Double-based parsing methods
    double ParseExpression();
    double ParseTerm();
    double ParseFactor();
    double ParsePower();

    MathValue ParseValueExpression();
    MathValue ParseValueTerm();
    MathValue ParseValueFactor();
    MathValue ParseValuePower();
    
    // Rational-based parsing methods
    Rational ParseExpressionRational();
    Rational ParseTermRational();
    Rational ParseFactorRational();
    Rational ParsePowerRational();
    
    void SkipSpace();
};

bool ParseLowerLimit(const std::wstring& s, std::wstring& var, double& val);
