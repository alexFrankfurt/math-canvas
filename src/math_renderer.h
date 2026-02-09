#pragma once

#include "math_types.h"
#include <windows.h>

class MathRenderer
{
public:
    static void Draw(HWND hEdit, HDC hdc, const MathObject& obj, size_t objIndex, const MathTypingState& state);
    static bool GetHitPart(HWND hEdit, HDC hdc, POINT ptMouse, size_t* outIndex, int* outPart);
    static COLORREF GetDefaultTextColor(HWND hEdit);
    static COLORREF GetActiveColor(HWND hEdit);
    static bool TryGetCharPos(HWND hEdit, LONG charIndex, POINT& outPt);
};
