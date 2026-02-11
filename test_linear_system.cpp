#include <iostream>
#include <vector>
#include <map>
#include "src/math_evaluator.h"

void PrintSolution(const std::map<std::wstring, double>& solution) {
    for (const auto& pair : solution) {
        if (pair.first == L"status") {
            std::wcout << L"Status: " << pair.second << std::endl;
            switch (static_cast<int>(pair.second)) {
                case 0: std::wcout << L"Success" << std::endl; break;
                case -1: std::wcout << L"Infinite solutions" << std::endl; break;
                case -2: std::wcout << L"No solution" << std::endl; break;
                case -3: std::wcout << L"No equations" << std::endl; break;
                case -4: std::wcout << L"Parse error" << std::endl; break;
                case -5: std::wcout << L"Underdetermined" << std::endl; break;
                case -6: std::wcout << L"Too many equations" << std::endl; break;
            }
        } else {
            std::wcout << pair.first << L" = " << pair.second << std::endl;
        }
    }
    std::wcout << std::endl;
}

int main() {
    MathEvaluator eval;
    
    // Test case 1: Simple 2x2 system
    std::wcout << L"Test 1: 2x + 3y = 7, 4x - y = 1" << std::endl;
    std::vector<std::wstring> equations1 = {L"2x+3y=7", L"4x-y=1"};
    auto result1 = eval.SolveSystemOfEquations(equations1);
    PrintSolution(result1);
    
    // Test case 2: System with zero coefficients
    std::wcout << L"Test 2: 2x = 6, 3y = 9" << std::endl;
    std::vector<std::wstring> equations2 = {L"2x=6", L"3y=9"};
    auto result2 = eval.SolveSystemOfEquations(equations2);
    PrintSolution(result2);
    
    // Test case 3: 3x3 system
    std::wcout << L"Test 3: x + y + z = 6, 2y + 5z = -4, 2x + 5y - z = 27" << std::endl;
    std::vector<std::wstring> equations3 = {L"x+y+z=6", L"2y+5z=-4", L"2x+5y-z=27"};
    auto result3 = eval.SolveSystemOfEquations(equations3);
    PrintSolution(result3);
    
    // Test case 4: Single equation
    std::wcout << L"Test 4: 3x = 12" << std::endl;
    std::vector<std::wstring> equations4 = {L"3x=12"};
    auto result4 = eval.SolveSystemOfEquations(equations4);
    PrintSolution(result4);
    
    // Test case 5: No solution
    std::wcout << L"Test 5: x + y = 1, x + y = 2" << std::endl;
    std::vector<std::wstring> equations5 = {L"x+y=1", L"x+y=2"};
    auto result5 = eval.SolveSystemOfEquations(equations5);
    PrintSolution(result5);
    
    return 0;
}