#include <iostream>
#include "src/math_evaluator.h"

int main() {
    MathEvaluator eval;
    double result = eval.Eval(L"2+2");
    std::cout << "2+2 = " << result << std::endl;
    return 0;
}