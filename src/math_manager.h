#pragma once

#include "math_evaluator.h"
#include "math_types.h"
#include <vector>
#include <string>

class MathManager
{
public:
    static MathManager& Get() { static MathManager instance; return instance; }

    std::vector<MathObject>& GetObjects() { return m_objects; }
/**
 * Get the current state of the math typing functionality
 * @return Reference to the current MathTypingState object
 */
    MathTypingState& GetState() { return m_state; } // Return reference to the math typing state

    void Clear() { m_objects.clear(); m_state = {}; }
    
    void ShiftObjectsAfter(LONG atPosInclusive, LONG delta);
    void DeleteObjectsInRange(LONG start, LONG end);
    bool IsPosInsideAnyObject(LONG pos, size_t* outIndex = nullptr);
    bool CanCalculateResult(const MathObject& obj) const;
    MathValue CalculateValueResult(const MathObject& obj) const;
    double CalculateResult(const MathObject& obj) const;
    std::wstring CalculateSystemResult(const MathObject& obj);
    std::wstring CalculateFormattedResult(const MathObject& obj) const;
    std::wstring FormatNumericResult(double value) const;
    std::wstring FormatValueResult(const MathValue& value) const;

private:
    MathManager() = default;
    std::vector<MathObject> m_objects;
    MathTypingState m_state;
};
