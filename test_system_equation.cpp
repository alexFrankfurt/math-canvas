#include <iostream>
#include <string>
#include <windows.h>
#include "src/math_manager.h"
#include "src/math_evaluator.h"
#include <locale>
#include <codecvt>

int main()
{
    // Use default locale
    
    // Test the system of equations solver
    MathManager& mgr = MathManager::Get();
    
    // Create a test system of equations object
    MathObject obj;
    obj.type = MathType::SystemOfEquations;
    obj.part3 = L"2x-14y=0, 8x+9y=0"; // Example from user request
    
    std::cout << "Testing system of equations solver..." << std::endl;
    std::wstring_convert<std::codecvt_utf8<wchar_t>> converter;
    std::cout << "Equations: " << converter.to_bytes(obj.part3) << std::endl;
    
    // Test the calculation directly
    MathEvaluator eval;
    std::vector<std::wstring> equations;
    equations.push_back(L"2x-14y=0");
    equations.push_back(L"8x+9y=0");
    auto solution = eval.SolveSystemOfEquations(equations);
    
    std::cout << "Direct solution for 2x-14y=0, 8x+9y=0:" << std::endl;
    for (const auto& [var, val] : solution) {
        std::cout << converter.to_bytes(var) << ": " << val << std::endl;
    }
    
    // Format the result manually
    std::wstring result = L" \uFF1D ";
    bool first = true;
    for (const auto& [var, val] : solution) {
        if (var == L"status") continue;
        
        if (!first) {
            result += L", ";
        }
        
        result += var + L"=";
        if (val == 0.0) {
            result += L"0";
        } else if (val == (long long)val) {
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
    
    std::cout << "Manually formatted result: " << converter.to_bytes(result) << std::endl;
    std::cout << "Result length: " << result.length() << std::endl;
    std::cout << "First character code: " << (int)result[0] << std::endl;
    
    return 0;
}