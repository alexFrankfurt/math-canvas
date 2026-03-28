#pragma once

#include "math_types.h"
#include <windows.h>

class MathRenderer
{
public:
    static void Draw(HWND hEdit, HDC hdc, const MathObject& obj, size_t objIndex, const MathTypingState& state);
    static bool TryGetObjectBounds(HWND hEdit, HDC hdc, const MathObject& obj, RECT& outRect);
    static bool TryGetActiveCaretPoint(HWND hEdit, HDC hdc, const MathObject& obj, size_t objIndex, const MathTypingState& state, POINT& outPt);
    static bool GetHitPart(HWND hEdit, HDC hdc, POINT ptMouse, size_t* outIndex, int* outPart, std::vector<size_t>* outNodePath = nullptr);
    static COLORREF GetDefaultTextColor(HWND hEdit);
    static COLORREF GetActiveColor(HWND hEdit);
    static bool TryGetCharPos(HWND hEdit, LONG charIndex, POINT& outPt);
};
