#pragma once

#include <windows.h>
#include <richedit.h>
#include <string>

// Public interface for the math/fraction engine
bool InstallMathSupport(HWND hRichEdit);
void ResetMathSupport();
void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator);
