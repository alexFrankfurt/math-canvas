#pragma once

#include "math_types.h"
#include <vector>
#include <string>

class MathManager
{
public:
    static MathManager& Get() { static MathManager instance; return instance; }

    std::vector<MathObject>& GetObjects() { return m_objects; }
    MathTypingState& GetState() { return m_state; }

    void Clear() { m_objects.clear(); m_state = {}; }
    
    void ShiftObjectsAfter(LONG atPosInclusive, LONG delta);
    void DeleteObjectsInRange(LONG start, LONG end);
    bool IsPosInsideAnyObject(LONG pos, size_t* outIndex = nullptr);
    double CalculateResult(const MathObject& obj);
    std::wstring CalculateSystemResult(const MathObject& obj);

private:
    MathManager() = default;
    std::vector<MathObject> m_objects;
    MathTypingState m_state;
};
