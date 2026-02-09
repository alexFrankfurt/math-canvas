#pragma once

#include <windows.h>
#include <string>
#include <vector>

enum class MathType { Fraction, Summation, Integral };

struct MathObject
{
    MathType type = MathType::Fraction;
    LONG barStart = 0;   // anchor character position
    LONG barLen = 0;     // anchor sequence length (5 for sum/int, variable for fraction)
    std::wstring part1;  // Numerator / Upper Limit
    std::wstring part2;  // Denominator / Lower Limit
    std::wstring part3;  // Expression / Function
    std::wstring resultText; // GDI-drawn result (e.g. "\uFF1D 302")
};

struct MathTypingState
{
    bool active = false;      
    int activePart = 0;       // 1 for part1, 2 for part2, 3 for part3
    size_t objectIndex = 0;   
};
