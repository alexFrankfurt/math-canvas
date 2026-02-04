#pragma once

#include <windows.h>
#include <richedit.h>
#include <string>

// Installs a subclass WndProc on the given RichEdit control.
// Keeps internal state needed to render stacked fractions (e.g., typing 3/4).
bool InstallFractionSupport(HWND hRichEdit);

// Clears any in-memory fraction state (call when clearing the editor text).
void ResetFractionSupport();

// Programmatic insertion (used by the "Insert Fraction" button).
void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator);
