#include <iostream>
#include <string>
#include <windows.h>
#include "src/math_manager.h"
#include "src/math_evaluator.h"
#include <locale>
#include <codecvt>
#include <vector>
#include <map>

// Helper function to format solution using CalculateSystemResult logic
std::wstring FormatSolution(const std::map<std::wstring, double>& solution) {
    std::wstring result = L" \uFF1D ";
    bool first = true;
    for (const auto& [var, val] : solution) {
        if (var == L"status") continue;
        
        if (!first) {
            result += L", ";
        }
        
        result += var + L"=";
        if (val == (long long)val) {
            result += std::to_wstring((long long)val);
        } else {
            wchar_t buf[64];
            swprintf(buf, 64, L"%.3f", val);
            result += buf;
        }
        
        first = false;
    }
    
    if (first) {
        result += L"x=0, y=0";
    }
    return result;
}

// Helper function to run a test case
void RunTestCase(const std::wstring& name, const std::vector<std::wstring>& equations, const std::wstring& expected_format) {
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    std::cout << "\n=== Test Case: " << converter.to_bytes(name) << " ===" << std::endl;
    
    MathEvaluator eval;
    auto solution = eval.SolveSystemOfEquations(equations);
    
    std::cout << "Equations: ";
    for (size_t i = 0; i < equations.size(); ++i) {
        if (i > 0) std::cout << ", ";
        std::cout << converter.to_bytes(equations[i]);
    }
    std::cout << std::endl;
    
    std::cout << "Solution status: " << solution[L"status"] << std::endl;
    
    std::wstring formatted = FormatSolution(solution);
    std::cout << "Formatted result: " << converter.to_bytes(formatted) << std::endl;
    
    if (!expected_format.empty()) {
        if (formatted == expected_format) {
            std::cout << "✓ Format matches expected" << std::endl;
        } else {
            std::cout << "✗ Format mismatch. Expected: " << converter.to_bytes(expected_format) << std::endl;
        }
    }
}

int main() {
    // Test Case 1: Trivial 2x2 system (zero solution)
    std::vector<std::wstring> eq1 = {L"2x-14y=0", L"8x+9y=0"};
    RunTestCase(L"Trivial Zero Solution", eq1, L" \uFF1D x=0, y=0");
    
    // Test Case2: Valid2x2 system (non-zero solution)
    std::vector<std::wstring> eq2 = {L"x+y=5", L"x-y=1"};
    RunTestCase(L"Valid 2x2 System", eq2, L" \uFF1D x=3, y=2");
    
    // Test Case3: Valid3x3 system
    std::vector<std::wstring> eq3 = {L"x+y+z=6", L"x-y+z=2", L"2x+y-z=1"};
    RunTestCase(L"Valid3x3 System", eq3, L" \uFF1D x=1, y=2, z=3");
    
    // Test Case4: Infinite solutions (2x2 dependent)
    std::vector<std::wstring> eq4 = {L"2x+2y=4", L"x+y=2"};
    RunTestCase(L"Infinite Solutions", eq4, L" \uFF1D x=0, y=0"); // Status -1, but format still shows x=0,y=0
    
    // Test Case5: No solution (2x2 inconsistent)
    std::vector<std::wstring> eq5 = {L"x+y=3", L"x+y=5"};
    RunTestCase(L"No Solution", eq5, L" \uFF1D No solution");
    
    // Test Case6: Single equation (underdetermined)
    std::vector<std::wstring> eq6 = {L"x+2y=5"};
    RunTestCase(L"Single Equation", eq6, L" \uFF1D "); // Status-5, no variables
    
    // Test Case7: Too many equations (4 equations)
    std::vector<std::wstring> eq7 = {L"x+y=5", L"x-y=1", L"2x+3y=13", L"3x-2y=4"};
    RunTestCase(L"Too Many Equations", eq7, L" \uFF1D Too many equations (max3)");
    
    // Test Case8: Decimal solution
    std::vector<std::wstring> eq8 = {L"2x+3y=10", L"4x-5y=1"};
    RunTestCase(L"Decimal Solution", eq8, L" \uFF1D x=2.500, y=1.667");
    
    std::cout << "\n=== All Tests Completed ===" << std::endl;
    return 0;
}