#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <windows.h>
#include <richedit.h>
#include <algorithm>
#include <iostream>
#include <string>
#include <vector>

#include "src/math_editor.h"
#include "src/math_manager.h"
#include "src/math_renderer.h"
#include "src/math_types.h"

namespace
{
    struct DocumentEntrySpec
    {
        size_t start = 0;
        size_t length = 0;
        std::wstring payload;
    };

    bool Check(bool condition, const std::wstring& label)
    {
        std::wcout << (condition ? L"[PASS] " : L"[FAIL] ") << label << std::endl;
        return condition;
    }

    void SendChars(HWND hwnd, const std::wstring& text)
    {
        for (wchar_t ch : text)
            SendMessage(hwnd, WM_CHAR, (WPARAM)ch, 0);
    }

    bool ContainsText(const std::vector<std::wstring>& values, const std::wstring& needle)
    {
        return std::find(values.begin(), values.end(), needle) != values.end();
    }

    HMODULE LoadRichEditWithFallback(const wchar_t** outClassName)
    {
        if (outClassName)
            *outClassName = nullptr;

        if (HMODULE mod = LoadLibraryW(L"Msftedit.dll"))
        {
            if (outClassName)
                *outClassName = L"RICHEDIT50W";
            return mod;
        }

        if (HMODULE mod = LoadLibraryW(L"Riched20.dll"))
        {
            if (outClassName)
                *outClassName = L"RICHEDIT20W";
            return mod;
        }

        return nullptr;
    }

    HWND CreateHiddenHostWindow(HINSTANCE instance)
    {
        const wchar_t className[] = L"WinDeskAppDocumentPersistenceTestWindow";
        WNDCLASSW wc = {};
        wc.lpfnWndProc = DefWindowProcW;
        wc.hInstance = instance;
        wc.lpszClassName = className;
        RegisterClassW(&wc);
        return CreateWindowExW(0, className, L"doc-test", WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT, CW_USEDEFAULT, 400, 300, nullptr, nullptr, instance, nullptr);
    }

    bool WriteUtf8File(const std::wstring& path, const std::wstring& text)
    {
        const int utf8Bytes = WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), nullptr, 0, nullptr, nullptr);
        if (utf8Bytes < 0)
            return false;

        std::vector<char> buffer((size_t)utf8Bytes);
        if (utf8Bytes > 0)
        {
            if (WideCharToMultiByte(CP_UTF8, 0, text.c_str(), (int)text.size(), buffer.data(), utf8Bytes, nullptr, nullptr) <= 0)
                return false;
        }

        HANDLE file = CreateFileW(path.c_str(), GENERIC_WRITE, 0, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE)
            return false;

        DWORD written = 0;
        const bool ok = buffer.empty() || WriteFile(file, buffer.data(), (DWORD)buffer.size(), &written, nullptr);
        CloseHandle(file);
        return ok && written == buffer.size();
    }

    bool ReadUtf8File(const std::wstring& path, std::wstring& outText)
    {
        outText.clear();

        HANDLE file = CreateFileW(path.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);
        if (file == INVALID_HANDLE_VALUE)
            return false;

        LARGE_INTEGER size = {};
        if (!GetFileSizeEx(file, &size) || size.QuadPart < 0 || size.QuadPart > MAXDWORD)
        {
            CloseHandle(file);
            return false;
        }

        std::vector<char> buffer((size_t)size.QuadPart);
        DWORD read = 0;
        const bool ok = buffer.empty() || ReadFile(file, buffer.data(), (DWORD)buffer.size(), &read, nullptr);
        CloseHandle(file);
        if (!ok || read != buffer.size())
            return false;

        size_t offset = 0;
        if (buffer.size() >= 3 && (unsigned char)buffer[0] == 0xEF && (unsigned char)buffer[1] == 0xBB && (unsigned char)buffer[2] == 0xBF)
            offset = 3;

        const int wideChars = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buffer.data() + offset, (int)(buffer.size() - offset), nullptr, 0);
        if (wideChars < 0)
            return false;

        outText.resize((size_t)wideChars);
        if (wideChars > 0)
        {
            if (MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, buffer.data() + offset, (int)(buffer.size() - offset), outText.data(), wideChars) <= 0)
                return false;
        }

        return true;
    }

    std::wstring BuildDocumentPayload(const std::wstring& rawText, const std::vector<DocumentEntrySpec>& entries, const std::wstring& version = L"D1")
    {
        std::wstring payload = version + L"|";
        MathObject::AppendString(payload, rawText);
        payload.push_back(L'|');
        MathObject::AppendCount(payload, entries.size());
        payload.push_back(L'[');
        for (const auto& entry : entries)
        {
            MathObject::AppendCount(payload, entry.start);
            payload.push_back(L'|');
            MathObject::AppendCount(payload, entry.length);
            payload.push_back(L'|');
            MathObject::AppendString(payload, entry.payload);
        }
        payload.push_back(L']');
        return payload;
    }
}

int wmain()
{
    int passed = 0;
    int failed = 0;
    auto run = [&](bool ok) { if (ok) ++passed; else ++failed; };

    const wchar_t* richEditClass = nullptr;
    HMODULE hRichEdit = LoadRichEditWithFallback(&richEditClass);
    run(Check(hRichEdit != nullptr && richEditClass != nullptr, L"load RichEdit library"));
    if (!hRichEdit || !richEditClass)
        return 1;

    HINSTANCE instance = GetModuleHandleW(nullptr);
    HWND host = CreateHiddenHostWindow(instance);
    run(Check(host != nullptr, L"create hidden host window"));
    if (!host)
    {
        FreeLibrary(hRichEdit);
        return 1;
    }

    HWND edit = CreateWindowExW(WS_EX_CLIENTEDGE, richEditClass, L"",
        WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN,
        0, 0, 300, 200, host, nullptr, instance, nullptr);
    run(Check(edit != nullptr, L"create RichEdit control"));
    if (!edit)
    {
        DestroyWindow(host);
        FreeLibrary(hRichEdit);
        return 1;
    }

    run(Check(InstallMathSupport(edit), L"install math editor support"));

    const auto resetEditor = [&]() {
        SetWindowTextW(edit, L"");
        ResetMathSupport();
        ApplyMathEditInsets(edit);
        SendMessage(edit, EM_SETSEL, 0, 0);
    };

    resetEditor();
    SendChars(edit, L"\\frac");
    SendMessage(edit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(edit, L"3 ");

    std::vector<std::wstring> unitSuggestions;
    size_t selectedSuggestion = 0;
    std::wstring suggestionPrefix;
    RECT popupRect = {};
    run(Check(DebugGetUnitSuggestionState(unitSuggestions, selectedSuggestion, &suggestionPrefix, &popupRect), L"unit dropdown appears after numeric space"));
    run(Check(suggestionPrefix.empty(), L"all-units dropdown starts unfiltered"));
    run(Check(ContainsText(unitSuggestions, L"m") && ContainsText(unitSuggestions, L"kg") && ContainsText(unitSuggestions, L"Pa"),
              L"all-units dropdown includes common symbols"));
    HDC popupHdc = GetDC(edit);
    RECT objectBounds = {};
    const bool hasObjectBounds = popupHdc != nullptr
        && !MathManager::Get().GetObjects().empty()
        && MathRenderer::TryGetObjectBounds(edit, popupHdc, MathManager::Get().GetObjects()[0], objectBounds);
    POINT firstAnchorPt = {};
    const bool hasAnchorPoint = !MathManager::Get().GetObjects().empty()
        && MathRenderer::TryGetCharPos(edit, MathManager::Get().GetObjects()[0].barStart, firstAnchorPt);
    if (popupHdc)
        ReleaseDC(edit, popupHdc);
    if (hasAnchorPoint)
        run(Check(firstAnchorPt.y >= 16, L"first-line math anchor keeps top clearance"));
    run(Check(popupRect.left >= 10 && popupRect.top >= 10, L"unit dropdown keeps top-left safety padding"));
    if (hasObjectBounds)
        run(Check(popupRect.left > objectBounds.left + 8, L"unit dropdown anchors to active caret instead of object edge"));

    SendChars(edit, L"c");
    unitSuggestions.clear();
    selectedSuggestion = 0;
    suggestionPrefix.clear();
    run(Check(DebugGetUnitSuggestionState(unitSuggestions, selectedSuggestion, &suggestionPrefix), L"unit dropdown stays visible while refining prefix"));
    run(Check(suggestionPrefix == L"c", L"unit dropdown tracks typed prefix"));
    run(Check(ContainsText(unitSuggestions, L"cm") && ContainsText(unitSuggestions, L"cd") && !ContainsText(unitSuggestions, L"kg"),
              L"typed unit prefix filters dropdown matches"));
    SendMessage(edit, WM_KEYDOWN, VK_RETURN, 0);
    run(Check(MathManager::Get().GetObjects().size() == 1 && MathManager::Get().GetObjects()[0].SlotText(1) == L"3 cm",
              L"accepting dropdown item inserts canonical unit text after spaced number"));
    unitSuggestions.clear();
    selectedSuggestion = 0;
    suggestionPrefix.clear();
    run(Check(!DebugGetUnitSuggestionState(unitSuggestions, selectedSuggestion, &suggestionPrefix), L"unit dropdown closes after accepting spaced suggestion"));

    resetEditor();
    SendChars(edit, L"\\frac");
    SendMessage(edit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(edit, L"3c");
    unitSuggestions.clear();
    selectedSuggestion = 0;
    suggestionPrefix.clear();
    run(Check(DebugGetUnitSuggestionState(unitSuggestions, selectedSuggestion, &suggestionPrefix), L"unit autocomplete appears for inline unit typing"));
    run(Check(ContainsText(unitSuggestions, L"cm") && ContainsText(unitSuggestions, L"cd"),
              L"inline unit typing offers matching completions"));
    SendMessage(edit, WM_KEYDOWN, VK_RETURN, 0);
    run(Check(MathManager::Get().GetObjects().size() == 1 && MathManager::Get().GetObjects()[0].SlotText(1) == L"3cm",
              L"accepting inline unit suggestion replaces typed prefix"));
    run(Check(MathManager::Get().GetState().active, L"accepting unit suggestion keeps active slot editing"));
    unitSuggestions.clear();
    selectedSuggestion = 0;
    suggestionPrefix.clear();
    run(Check(!DebugGetUnitSuggestionState(unitSuggestions, selectedSuggestion, &suggestionPrefix), L"unit dropdown closes after accepting inline suggestion"));
    SendMessage(edit, WM_CHAR, (WPARAM)L'*', 0);
    unitSuggestions.clear();
    selectedSuggestion = 0;
    suggestionPrefix.clear();
    run(Check(DebugGetUnitSuggestionState(unitSuggestions, selectedSuggestion, &suggestionPrefix), L"unit dropdown reopens after unit operator"));
    run(Check(suggestionPrefix.empty() && ContainsText(unitSuggestions, L"s"), L"operator-triggered dropdown shows full unit list"));

    resetEditor();
    SendChars(edit, L"A ");
    SendChars(edit, L"\\frac");
    SendMessage(edit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(edit, L"3m");
    SendMessage(edit, WM_KEYDOWN, VK_TAB, 0);
    SendChars(edit, L"4s");
    SendMessage(edit, WM_KEYDOWN, VK_RETURN, 0);
    SendChars(edit, L" B ");
    SendChars(edit, L"\\sqrt");
    SendMessage(edit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(edit, L"9+\\frac");
    SendMessage(edit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(edit, L"16");
    SendMessage(edit, WM_KEYDOWN, VK_TAB, 0);
    SendChars(edit, L"4");
    SendMessage(edit, WM_KEYDOWN, VK_RETURN, 0);
    SendChars(edit, L" C");

    std::wstring originalPayload;
    run(Check(SerializeMathDocument(edit, originalPayload), L"serialize structured document payload"));
    run(Check(originalPayload.rfind(L"D1|", 0) == 0, L"document payload uses versioned D1 prefix"));
    run(Check(MathManager::Get().GetObjects().size() == 2, L"document contains two structured objects before reload"));

    SetWindowTextW(edit, L"stomped");
    ResetMathSupport();
    run(Check(TryDeserializeMathDocument(edit, originalPayload), L"reload structured document payload"));

    std::wstring roundTripPayload;
    run(Check(SerializeMathDocument(edit, roundTripPayload), L"reserialize document after reload"));
    run(Check(roundTripPayload == originalPayload, L"document payload round-trips exactly"));
    run(Check(MathManager::Get().GetObjects().size() == 2, L"document contains two structured objects after reload"));
    if (MathManager::Get().GetObjects().size() == 2)
    {
        const auto& firstObj = MathManager::Get().GetObjects()[0];
        const auto& secondObj = MathManager::Get().GetObjects()[1];
        run(Check(firstObj.type == MathType::Fraction, L"round-tripped first object type preserved"));
        run(Check(firstObj.SlotText(1) == L"3m", L"round-tripped numerator with units preserved"));
        run(Check(firstObj.SlotText(2) == L"4s", L"round-tripped denominator with units preserved"));
        run(Check(secondObj.type == MathType::SquareRoot, L"round-tripped second object type preserved"));
        run(Check(secondObj.SlotText(1).find(L"((16)/(4))") != std::wstring::npos,
                  L"round-tripped nested fraction inside square root preserved"));
    }

    const int lenAfterReload = GetWindowTextLengthW(edit);
    std::vector<wchar_t> buffer((size_t)lenAfterReload + 1, L'\0');
    GetWindowTextW(edit, buffer.data(), lenAfterReload + 1);
    const std::wstring reloadedText = std::wstring(buffer.data());
    run(Check(reloadedText.find(L"A ") == 0, L"plain text before structured object preserved"));
    run(Check(reloadedText.find(L"B") != std::wstring::npos, L"plain text between structured objects preserved"));
    run(Check(reloadedText.find(L"C") != std::wstring::npos, L"plain text after structured objects preserved"));

    wchar_t tempPath[MAX_PATH] = L"";
    wchar_t tempFile[MAX_PATH] = L"";
    run(Check(GetTempPathW(MAX_PATH, tempPath) != 0 && GetTempFileNameW(tempPath, L"wdm", 0, tempFile) != 0,
              L"create temp file for file-backed document round-trip"));
    if (tempFile[0] != L'\0')
    {
        run(Check(WriteUtf8File(tempFile, originalPayload), L"write multi-object document payload to disk"));
        std::wstring filePayload;
        run(Check(ReadUtf8File(tempFile, filePayload), L"read multi-object document payload from disk"));
        run(Check(filePayload == originalPayload, L"file-backed payload bytes round-trip exactly"));

        SetWindowTextW(edit, L"mutated");
        ResetMathSupport();
        run(Check(TryDeserializeMathDocument(edit, filePayload), L"reload file-backed multi-object document payload"));
        run(Check(MathManager::Get().GetObjects().size() == 2, L"file-backed reload preserves two structured objects"));
        DeleteFileW(tempFile);
    }

    run(Check(!TryDeserializeMathDocument(edit, L"not-a-document"), L"invalid document prefix is rejected"));
    run(Check(!TryDeserializeMathDocument(edit, L"D2|" + originalPayload.substr(3)), L"unknown document version prefix is rejected"));
    run(Check(!TryDeserializeMathDocument(edit, originalPayload.substr(0, originalPayload.size() - 1)), L"truncated document payload is rejected"));
    run(Check(!TryDeserializeMathDocument(edit, originalPayload + L"extra"), L"document payload with trailing garbage is rejected"));

    const auto& objects = MathManager::Get().GetObjects();
    if (objects.size() == 2)
    {
        const size_t firstStart = (size_t)objects[0].barStart;
        const size_t firstLen = (size_t)objects[0].barLen;
        const std::wstring firstPayload = objects[0].SerializeTransferPayload();
        const size_t secondStart = (size_t)objects[1].barStart;
        const size_t secondLen = (size_t)objects[1].barLen;
        const std::wstring secondPayload = objects[1].SerializeTransferPayload();

        const std::wstring invalidObjectPayload = BuildDocumentPayload(
            reloadedText,
            {
                { firstStart, firstLen, L"not-a-math-object" },
                { secondStart, secondLen, secondPayload }
            });
        run(Check(!TryDeserializeMathDocument(edit, invalidObjectPayload), L"document with invalid embedded object payload is rejected"));

        const std::wstring overlappingEntriesPayload = BuildDocumentPayload(
            reloadedText,
            {
                { firstStart, firstLen, firstPayload },
                { firstStart, secondLen, secondPayload }
            });
        run(Check(!TryDeserializeMathDocument(edit, overlappingEntriesPayload), L"document with overlapping entries is rejected"));

        const std::wstring outOfRangeEntryPayload = BuildDocumentPayload(
            reloadedText,
            {
                { (size_t)reloadedText.size() + 1, 3, firstPayload }
            });
        run(Check(!TryDeserializeMathDocument(edit, outOfRangeEntryPayload), L"document with out-of-range entry is rejected"));

        const std::wstring zeroLengthEntryPayload = BuildDocumentPayload(
            reloadedText,
            {
                { firstStart, 0, firstPayload }
            });
        run(Check(!TryDeserializeMathDocument(edit, zeroLengthEntryPayload), L"document with zero-length entry is rejected"));
    }

    DestroyWindow(edit);
    DestroyWindow(host);
    FreeLibrary(hRichEdit);

    std::wcout << L"\n=== Summary ===" << std::endl;
    std::wcout << L"Passed: " << passed << std::endl;
    std::wcout << L"Failed: " << failed << std::endl;
    return failed == 0 ? 0 : 1;
}