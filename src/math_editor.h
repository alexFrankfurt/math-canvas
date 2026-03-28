#pragma once

#include <windows.h>
#include <richedit.h>
#include <string>
#include <vector>

// Public interface for the math/fraction engine
bool InstallMathSupport(HWND hRichEdit);
void ResetMathSupport();
void ApplyMathEditInsets(HWND hRichEdit);
void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator);
bool SerializeMathDocument(HWND hEdit, std::wstring& outPayload);
bool TryDeserializeMathDocument(HWND hEdit, const std::wstring& payload);
bool DebugRunStructuredRoundTripSelfTest(HWND hEdit, std::wstring& outDetails);
bool DebugRunStructuredFragmentRoundTripSelfTest(HWND hEdit, std::wstring& outDetails);
bool DebugRunStructuredDocumentRoundTripSelfTest(HWND hEdit, std::wstring& outDetails);
bool DebugGetUnitSuggestionState(std::vector<std::wstring>& outSuggestions, size_t& outSelectedIndex, std::wstring* outPrefix = nullptr, RECT* outPopupRect = nullptr);
