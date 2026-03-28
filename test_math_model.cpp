#include <cmath>
#include <iostream>
#include <string>

#include "src/math_manager.h"
#include "src/math_types.h"
#include "src/math_evaluator.h"

namespace {
    constexpr double kEps = 1e-6;

    bool NearlyEqual(double a, double b)
    {
        return std::fabs(a - b) <= kEps;
    }

    bool Check(bool condition, const std::wstring& label)
    {
        std::wcout << (condition ? L"[PASS] " : L"[FAIL] ") << label << std::endl;
        return condition;
    }

    bool CheckNear(double actual, double expected, const std::wstring& label)
    {
        const bool ok = NearlyEqual(actual, expected);
        std::wcout << (ok ? L"[PASS] " : L"[FAIL] ")
                   << label << L" | expected=" << expected
                   << L" | actual=" << actual << std::endl;
        return ok;
    }
}

int main()
{
    int passed = 0;
    int failed = 0;
    auto run = [&](bool ok) { if (ok) ++passed; else ++failed; };

    MathEvaluator eval;

    MathObject sqrtObj;
    sqrtObj.type = MathType::SquareRoot;
    sqrtObj.SetParts();
    sqrtObj.EnsureStructuredEditLeaf(1);
    sqrtObj.EditableLeafText(1) = L"9+\\frac";
    std::vector<size_t> nestedFractionPath;
    run(Check(sqrtObj.InsertNestedNode(1, {}, L"\\frac", MathNodeKind::Fraction, nestedFractionPath, 0),
              L"insert nested fraction into square root"));
    sqrtObj.EditableLeafText(1, &nestedFractionPath) = L"16";
    std::vector<size_t> denominatorPath = nestedFractionPath;
    run(Check(sqrtObj.MoveToSiblingSlot(1, denominatorPath, 1), L"move fraction path to denominator"));
    sqrtObj.EditableLeafText(1, &denominatorPath) = L"4";
    sqrtObj.SyncLegacyFromSlots();
    run(Check(sqrtObj.SlotText(1).find(L"((16)/(4))") != std::wstring::npos,
              L"fraction flatten text present in radicand"));
    run(CheckNear(eval.Eval(sqrtObj.SlotText(1)), 13.0, L"flattened nested fraction evaluates"));

    MathObject powerObj;
    powerObj.type = MathType::Power;
    powerObj.SetParts();
    powerObj.EnsureStructuredEditLeaf(1);
    powerObj.EnsureStructuredEditLeaf(2);
    powerObj.EditableLeafText(1) = L"2";
    powerObj.EditableLeafText(2) = L"3";
    powerObj.SyncLegacyFromSlots();
    run(CheckNear(eval.Eval(powerObj.SlotText(1) + L"^" + powerObj.SlotText(2)), 8.0, L"power slots still evaluate"));

    MathObject absObj;
    absObj.type = MathType::AbsoluteValue;
    absObj.SetParts();
    absObj.EnsureStructuredEditLeaf(1);
    absObj.EditableLeafText(1) = L"-5+\\pow";
    std::vector<size_t> nestedPowerPath;
    run(Check(absObj.InsertNestedNode(1, {}, L"\\pow", MathNodeKind::Power, nestedPowerPath, 0),
              L"insert nested power into absolute value"));
    absObj.EditableLeafText(1, &nestedPowerPath) = L"2";
    std::vector<size_t> exponentPath = nestedPowerPath;
    run(Check(absObj.MoveToSiblingSlot(1, exponentPath, 1), L"move power path to exponent"));
    absObj.EditableLeafText(1, &exponentPath) = L"4";
    absObj.SyncLegacyFromSlots();
    run(Check(absObj.SlotText(1).find(L"((2)^(4))") != std::wstring::npos,
              L"power flatten text present in absolute value"));
    run(CheckNear(eval.Eval(L"abs(" + absObj.SlotText(1) + L")"), 11.0, L"absolute value with nested power evaluates"));

    std::vector<MathNode> nodes;
    nodes.push_back(MathNode::MakeText(L"a"));
    nodes.push_back(MathNode::MakeText(L"b"));
    MathObject::NormalizeNodeSequence(nodes);
    run(Check(nodes.size() == 1 && nodes[0].kind == MathNodeKind::Text && nodes[0].text == L"ab",
              L"normalize merges adjacent text nodes"));

    MathObject determinantObj;
    determinantObj.type = MathType::Determinant;
    determinantObj.SetMatrix2x2(L"1+1", L"3", L"4", L"5");
    determinantObj.EnsureStructuredEditLeaf(1);
    determinantObj.EnsureStructuredEditLeaf(2);
    determinantObj.EnsureStructuredEditLeaf(3);
    determinantObj.EnsureStructuredEditLeaf(4);
    determinantObj.SyncLegacyFromSlots();
    MathManager::Get().Clear();
    run(CheckNear(MathManager::Get().CalculateResult(determinantObj), -2.0,
                  L"determinant uses structured 2x2 cell slots"));

    MathObject matrixObj;
    matrixObj.type = MathType::Matrix;
    matrixObj.SetMatrix2x2(L"a", L"b", L"c", L"d");
    matrixObj.EnsureStructuredEditLeaf(1);
    matrixObj.EnsureStructuredEditLeaf(2);
    matrixObj.EnsureStructuredEditLeaf(3);
    matrixObj.EnsureStructuredEditLeaf(4);
    matrixObj.EditableLeafText(4) = L"7";
    matrixObj.SyncLegacyFromSlots();
    run(Check(matrixObj.SlotText(4) == L"7", L"matrix fourth cell stays slot-backed"));

    MathObject serializedSqrtObj = sqrtObj;
    serializedSqrtObj.resultText = L" \uFF1D 13";
    const std::wstring serializedSqrtPayload = serializedSqrtObj.SerializeTransferPayload();
    MathObject deserializedSqrtObj;
    run(Check(MathObject::TryDeserializeTransferPayload(serializedSqrtPayload, deserializedSqrtObj),
              L"deserialize nested sqrt transfer payload"));
    run(Check(deserializedSqrtObj.type == MathType::SquareRoot, L"round-tripped sqrt type preserved"));
    run(Check(deserializedSqrtObj.SlotText(1) == serializedSqrtObj.SlotText(1),
              L"round-tripped sqrt radicand text preserved"));
    run(Check(deserializedSqrtObj.resultText == serializedSqrtObj.resultText,
              L"round-tripped result text preserved"));
    run(Check(!deserializedSqrtObj.CloneStructuredSlotNodes(1).empty() && deserializedSqrtObj.CloneStructuredSlotNodes(1)[1].kind == MathNodeKind::Fraction,
              L"round-tripped nested fraction node preserved"));

    MathObject serializedMatrixObj = matrixObj;
    const std::wstring serializedMatrixPayload = serializedMatrixObj.SerializeTransferPayload();
    MathObject deserializedMatrixObj;
    run(Check(MathObject::TryDeserializeTransferPayload(serializedMatrixPayload, deserializedMatrixObj),
              L"deserialize matrix transfer payload"));
    run(Check(deserializedMatrixObj.type == MathType::Matrix, L"round-tripped matrix type preserved"));
    run(Check(deserializedMatrixObj.SlotText(4) == L"7", L"round-tripped matrix cell text preserved"));
    run(Check(deserializedMatrixObj.slots.size() == 4, L"round-tripped matrix slot count preserved"));

    run(Check(serializedSqrtObj.BuildPlainTextFallback() == L"sqrt(9+((16)/(4)))",
              L"square root plain-text fallback is canonical"));

    MathObject fallbackFractionObj;
    fallbackFractionObj.type = MathType::Fraction;
    fallbackFractionObj.SetParts(L"x+1", L"y-1");
    run(Check(fallbackFractionObj.BuildPlainTextFallback() == L"(x+1)/(y-1)",
              L"fraction plain-text fallback is canonical"));

    MathObject slotBackedObj;
    slotBackedObj.type = MathType::Fraction;
    slotBackedObj.SetParts(L"slot-num", L"slot-den");
    slotBackedObj.part1 = L"legacy-num";
    slotBackedObj.part2 = L"legacy-den";
    run(Check(slotBackedObj.PartText(1) == L"slot-num" && slotBackedObj.PartText(2) == L"slot-den",
              L"PartText prefers slot-backed values over stale legacy mirrors"));
    run(Check(slotBackedObj.BuildPlainTextFallback() == L"(slot-num)/(slot-den)",
              L"plain-text fallback uses slot-backed values over stale legacy mirrors"));

    MathObject mirrorRefreshObj;
    mirrorRefreshObj.type = MathType::Fraction;
    mirrorRefreshObj.SetParts(L"a", L"b");
    mirrorRefreshObj.SetPartText(1, L"updated-num");
    run(Check(mirrorRefreshObj.part1 == L"updated-num" && mirrorRefreshObj.part2 == L"b",
              L"SetPartText refreshes only touched legacy mirrors"));

    MathObject matrixMirrorObj;
    matrixMirrorObj.type = MathType::Matrix;
    matrixMirrorObj.SetMatrix2x2(L"w", L"x", L"y", L"z");
    run(Check(matrixMirrorObj.part1 == L"w" && matrixMirrorObj.part2 == L"x" && matrixMirrorObj.part3 == L"y",
              L"SetMatrix2x2 keeps first three legacy mirrors aligned"));

    MathObject fallbackLogObj;
    fallbackLogObj.type = MathType::Logarithm;
    fallbackLogObj.SetParts(L"2", L"8");
    run(Check(fallbackLogObj.BuildPlainTextFallback() == L"log_[2](8)",
              L"log plain-text fallback is canonical"));

    MathObject fallbackDetObj;
    fallbackDetObj.type = MathType::Determinant;
    fallbackDetObj.SetMatrix2x2(L"1", L"2", L"3", L"4");
    run(Check(fallbackDetObj.BuildPlainTextFallback() == L"det([[1, 2], [3, 4]])",
              L"determinant plain-text fallback is canonical"));

    MathObject fallbackSystemObj;
    fallbackSystemObj.type = MathType::SystemOfEquations;
    fallbackSystemObj.SetParts(L"x+y=3", L"x-y=1", L"");
    run(Check(fallbackSystemObj.BuildPlainTextFallback() == L"x+y=3\r\nx-y=1",
              L"system plain-text fallback preserves line breaks"));

    MathObject invalidLogObj;
    invalidLogObj.type = MathType::Logarithm;
    invalidLogObj.SetParts(L"1", L"8");
    run(Check(MathManager::Get().CalculateFormattedResult(invalidLogObj) == L" \uFF1D invalid log base",
              L"invalid log base surfaces explicit result text"));

    MathObject incompleteFractionObj;
    incompleteFractionObj.type = MathType::Fraction;
    incompleteFractionObj.SetParts(L"5", L"");
    run(Check(MathManager::Get().CalculateFormattedResult(incompleteFractionObj) == L" \uFF1D incomplete",
              L"incomplete fraction surfaces explicit result text"));

    MathObject zeroDenominatorObj;
    zeroDenominatorObj.type = MathType::Fraction;
    zeroDenominatorObj.SetParts(L"5", L"2-2");
    run(Check(MathManager::Get().CalculateFormattedResult(zeroDenominatorObj) == L" \uFF1D undefined",
              L"zero denominator surfaces undefined result text"));

    MathObject unitSumObj;
    unitSumObj.type = MathType::Sum;
    unitSumObj.SetParts(L"3m + 40cm");
    run(Check(MathManager::Get().CalculateFormattedResult(unitSumObj) == L" \uFF1D 3.4 m",
              L"sum object formats compatible unit addition"));

    MathObject unitFractionObj;
    unitFractionObj.type = MathType::Fraction;
    unitFractionObj.SetParts(L"10m", L"2s");
    run(Check(MathManager::Get().CalculateFormattedResult(unitFractionObj) == L" \uFF1D 5 m/s",
              L"fraction object formats composed units"));

    MathObject unitSqrtObj;
    unitSqrtObj.type = MathType::SquareRoot;
    unitSqrtObj.SetParts(L"9m^2", L"2");
    run(Check(MathManager::Get().CalculateFormattedResult(unitSqrtObj) == L" \uFF1D 3 m",
              L"square root object reduces even unit powers"));

    MathObject incompatibleUnitObj;
    incompatibleUnitObj.type = MathType::Sum;
    incompatibleUnitObj.SetParts(L"3m + 2s");
    run(Check(MathManager::Get().CalculateFormattedResult(incompatibleUnitObj) == L" \uFF1D incompatible units",
              L"incompatible unit arithmetic surfaces explicit result text"));

    MathObject invalidUnitExponentObj;
    invalidUnitExponentObj.type = MathType::SquareRoot;
    invalidUnitExponentObj.SetParts(L"9m", L"2");
    run(Check(MathManager::Get().CalculateFormattedResult(invalidUnitExponentObj) == L" \uFF1D invalid unit exponent",
              L"invalid unit exponent surfaces explicit result text"));

    MathObject unitLogObj;
    unitLogObj.type = MathType::Logarithm;
    unitLogObj.SetParts(L"10", L"10m");
    run(Check(MathManager::Get().CalculateFormattedResult(unitLogObj) == L" \uFF1D log requires abstract number",
              L"logarithm rejects dimensional arguments with explicit result text"));

    std::wcout << L"\n=== Summary ===" << std::endl;
    std::wcout << L"Passed: " << passed << std::endl;
    std::wcout << L"Failed: " << failed << std::endl;
    return failed == 0 ? 0 : 1;
}