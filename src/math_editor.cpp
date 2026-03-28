#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "math_editor.h"
#include "math_manager.h"
#include "math_renderer.h"
#include <cwctype>
#include <algorithm>

    // Make anchor characters invisible by setting their text color to the
    // background color.  This prevents RichEdit from drawing U+2500 glyphs;
    // the renderer draws its own bar via GDI instead.
    static void HideAnchorChars(HWND hwnd, LONG start, LONG len)
    {
        DWORD oldS, oldE;
        SendMessage(hwnd, EM_GETSEL, (WPARAM)&oldS, (LPARAM)&oldE);
        SendMessage(hwnd, EM_SETSEL, (WPARAM)start, (LPARAM)(start + len));

        COLORREF bk = (COLORREF)SendMessage(hwnd, EM_SETBKGNDCOLOR, 0, 0);
        SendMessage(hwnd, EM_SETBKGNDCOLOR, 0, (LPARAM)bk); // restore

        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_COLOR;
        cf.crTextColor = bk;
        cf.dwEffects = 0; // clear CFE_AUTOCOLOR
        SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);

        SendMessage(hwnd, EM_SETSEL, (WPARAM)oldS, (LPARAM)oldE);
    }

    static bool BuildMathDirtyRect(HWND hwnd, RECT& outRc);

    static void RequestMathRepaint(HWND hwnd)
    {
        RECT dirty = {};
        if (BuildMathDirtyRect(hwnd, dirty))
            RedrawWindow(hwnd, &dirty, nullptr, RDW_INVALIDATE | RDW_NOERASE);
        else
            RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
    }

    static void InvalidateMathOverlay(HWND hwnd)
    {
        RedrawWindow(hwnd, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
    }

    static void RestoreTypingFormat(HWND hwnd)
    {
        COLORREF bk = (COLORREF)SendMessage(hwnd, EM_SETBKGNDCOLOR, 0, 0);
        SendMessage(hwnd, EM_SETBKGNDCOLOR, 0, (LPARAM)bk);
        const int brightness = (GetRValue(bk) * 299 + GetGValue(bk) * 587 + GetBValue(bk) * 114) / 1000;
        const COLORREF textColor = (brightness < 128) ? RGB(220, 220, 220) : RGB(0, 0, 0);

        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_COLOR;
        cf.dwEffects = 0;
        cf.crTextColor = textColor;
        SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
    }

    static bool HasObjectAtOrAfter(const std::vector<MathObject>& objects, LONG atPosInclusive)
    {
        for (const auto& obj : objects)
            if (obj.barStart >= atPosInclusive) return true;
        return false;
    }

    static bool BuildMathDirtyRect(HWND hwnd, RECT& outRc)
    {
        auto& objects = MathManager::Get().GetObjects();
        if (objects.empty()) return false;

        HDC hdc = GetDC(hwnd);
        if (!hdc) return false;

        HFONT baseFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);
        if (!baseFont) baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT oldFont = (HFONT)SelectObject(hdc, baseFont);
        TEXTMETRICW tm = {};
        GetTextMetricsW(hdc, &tm);

        bool hasRect = false;
        RECT unionRc = {};
        for (const auto& obj : objects)
        {
            RECT rc = {};
            if (!MathRenderer::TryGetObjectBounds(hwnd, hdc, obj, rc))
                continue;

            InflateRect(&rc, tm.tmAveCharWidth * 2, tm.tmHeight);

            if (!hasRect) {
                unionRc = rc;
                hasRect = true;
            }
            else {
                unionRc.left = (std::min)(unionRc.left, rc.left);
                unionRc.top = (std::min)(unionRc.top, rc.top);
                unionRc.right = (std::max)(unionRc.right, rc.right);
                unionRc.bottom = (std::max)(unionRc.bottom, rc.bottom);
            }
        }

        SelectObject(hdc, oldFont);
        ReleaseDC(hwnd, hdc);

        if (!hasRect) return false;

        RECT client = {};
        GetClientRect(hwnd, &client);
        if (!IntersectRect(&outRc, &unionRc, &client)) return false;
        return true;
    }

    struct ScopedNoRedraw
    {
        HWND hwnd;
        explicit ScopedNoRedraw(HWND h) : hwnd(h)
        {
            SendMessage(hwnd, WM_SETREDRAW, FALSE, 0);
        }
        ~ScopedNoRedraw()
        {
            SendMessage(hwnd, WM_SETREDRAW, TRUE, 0);
        }
    };

    struct EquationResult {
        double value;
        std::wstring message;
        bool success;

        EquationResult() : value(0), success(false) {}
        EquationResult(double v, const std::wstring& msg, bool s) : value(v), message(msg), success(s) {}
    };

namespace
{
    constexpr LONG kMathEditInsetLeft = 8;
    constexpr LONG kMathEditInsetTop = 24;
    constexpr LONG kMathEditInsetRight = 8;
    constexpr LONG kMathEditInsetBottom = 6;

    HWND g_hEdit = nullptr;
    WNDPROC g_originalProc = nullptr;
    std::wstring g_currentCommand;
    std::wstring g_currentNumber;
    bool g_suppressNextChar = false;

    struct UnitSuggestionContext
    {
        size_t replaceStart = 0;
        std::wstring prefix;
        bool showAll = false;
    };

    struct UnitSuggestionPopupState
    {
        bool visible = false;
        size_t replaceStart = 0;
        size_t selectedIndex = 0;
        size_t topIndex = 0;
        std::wstring prefix;
        RECT popupRect = {};
        std::vector<std::wstring> items;
        std::vector<RECT> itemRects;
    };

    UnitSuggestionPopupState g_unitSuggestionPopup;

    static UINT GetMathClipboardFormat()
    {
        static UINT format = RegisterClipboardFormatW(L"WinDeskApp.MathObjectPayload");
        return format;
    }

    struct MathClipboardFragmentEntry
    {
        LONG relativeStart = 0;
        std::wstring payload;
    };

    struct MathClipboardFragment
    {
        std::wstring rawText;
        std::vector<MathClipboardFragmentEntry> entries;
    };

    struct MathDocumentEntry
    {
        LONG start = 0;
        LONG length = 0;
        std::wstring payload;
    };

    struct MathDocumentSnapshot
    {
        std::wstring rawText;
        std::vector<MathDocumentEntry> entries;
    };

    static UINT GetMathFragmentClipboardFormat()
    {
        static UINT format = RegisterClipboardFormatW(L"WinDeskApp.MathFragmentPayload");
        return format;
    }

    static std::wstring ExtractTextRange(HWND hwnd, LONG start, LONG end)
    {
        if (end <= start)
            return L"";

        std::vector<wchar_t> buffer((size_t)(end - start) + 1, L'\0');
        TEXTRANGEW range = { { start, end }, buffer.data() };
        SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&range);
        return std::wstring(buffer.data());
    }

    static bool SetClipboardUnicodeText(const std::wstring& text)
    {
        const size_t bytes = (text.size() + 1) * sizeof(wchar_t);
        HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (!memory)
            return false;

        void* locked = GlobalLock(memory);
        if (!locked)
        {
            GlobalFree(memory);
            return false;
        }

        memcpy(locked, text.c_str(), bytes);
        GlobalUnlock(memory);
        if (!SetClipboardData(CF_UNICODETEXT, memory))
        {
            GlobalFree(memory);
            return false;
        }
        return true;
    }

    static bool SetClipboardPayload(UINT format, const std::wstring& payload)
    {
        const size_t bytes = (payload.size() + 1) * sizeof(wchar_t);
        HGLOBAL memory = GlobalAlloc(GMEM_MOVEABLE, bytes);
        if (!memory)
            return false;

        void* locked = GlobalLock(memory);
        if (!locked)
        {
            GlobalFree(memory);
            return false;
        }

        memcpy(locked, payload.c_str(), bytes);
        GlobalUnlock(memory);
        if (!SetClipboardData(format, memory))
        {
            GlobalFree(memory);
            return false;
        }
        return true;
    }

    static bool ReadClipboardPayload(UINT format, std::wstring& out)
    {
        HANDLE handle = GetClipboardData(format);
        if (!handle)
            return false;

        const wchar_t* text = static_cast<const wchar_t*>(GlobalLock(handle));
        if (!text)
            return false;
        out = text;
        GlobalUnlock(handle);
        return true;
    }

    static std::wstring SerializeClipboardFragment(const MathClipboardFragment& fragment)
    {
        std::wstring payload = L"F1|";
        MathObject::AppendString(payload, fragment.rawText);
        payload.push_back(L'|');
        MathObject::AppendCount(payload, fragment.entries.size());
        payload.push_back(L'[');
        for (const auto& entry : fragment.entries)
        {
            MathObject::AppendCount(payload, (size_t)entry.relativeStart);
            payload.push_back(L'|');
            MathObject::AppendString(payload, entry.payload);
        }
        payload.push_back(L']');
        return payload;
    }

    static std::wstring SerializeMathDocumentSnapshot(const MathDocumentSnapshot& snapshot)
    {
        std::wstring payload = L"D1|";
        MathObject::AppendString(payload, snapshot.rawText);
        payload.push_back(L'|');
        MathObject::AppendCount(payload, snapshot.entries.size());
        payload.push_back(L'[');
        for (const auto& entry : snapshot.entries)
        {
            MathObject::AppendCount(payload, (size_t)entry.start);
            payload.push_back(L'|');
            MathObject::AppendCount(payload, (size_t)entry.length);
            payload.push_back(L'|');
            MathObject::AppendString(payload, entry.payload);
        }
        payload.push_back(L']');
        return payload;
    }

    static bool TryDeserializeMathDocumentSnapshot(const std::wstring& payload, MathDocumentSnapshot& snapshot)
    {
        if (payload.size() < 3 || payload.substr(0, 3) != L"D1|")
            return false;

        size_t cursor = 3;
        if (!MathObject::ParseString(payload, cursor, snapshot.rawText))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'|')
            return false;
        ++cursor;

        size_t count = 0;
        if (!MathObject::ParseCount(payload, cursor, count))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'[')
            return false;
        ++cursor;

        snapshot.entries.clear();
        snapshot.entries.reserve(count);
        for (size_t entryIndex = 0; entryIndex < count; ++entryIndex)
        {
            size_t start = 0;
            size_t length = 0;
            if (!MathObject::ParseCount(payload, cursor, start))
                return false;
            if (cursor >= payload.size() || payload[cursor] != L'|')
                return false;
            ++cursor;
            if (!MathObject::ParseCount(payload, cursor, length))
                return false;
            if (cursor >= payload.size() || payload[cursor] != L'|')
                return false;
            ++cursor;

            MathDocumentEntry entry;
            entry.start = (LONG)start;
            entry.length = (LONG)length;
            if (!MathObject::ParseString(payload, cursor, entry.payload))
                return false;
            snapshot.entries.push_back(std::move(entry));
        }

        if (cursor >= payload.size() || payload[cursor] != L']')
            return false;
        ++cursor;
        if (cursor != payload.size())
            return false;

        LONG previousEnd = -1;
        for (const auto& entry : snapshot.entries)
        {
            if (entry.start < 0 || entry.length <= 0)
                return false;
            if (entry.start + entry.length > (LONG)snapshot.rawText.size())
                return false;
            if (previousEnd > entry.start)
                return false;
            previousEnd = entry.start + entry.length;
        }

        return true;
    }

    static bool TryDeserializeClipboardFragment(const std::wstring& payload, MathClipboardFragment& fragment)
    {
        if (payload.size() < 3 || payload.substr(0, 3) != L"F1|")
            return false;

        size_t cursor = 3;
        if (!MathObject::ParseString(payload, cursor, fragment.rawText))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'|')
            return false;
        ++cursor;

        size_t count = 0;
        if (!MathObject::ParseCount(payload, cursor, count))
            return false;
        if (cursor >= payload.size() || payload[cursor] != L'[')
            return false;
        ++cursor;

        fragment.entries.clear();
        fragment.entries.reserve(count);
        for (size_t entryIndex = 0; entryIndex < count; ++entryIndex)
        {
            size_t relativeStart = 0;
            if (!MathObject::ParseCount(payload, cursor, relativeStart))
                return false;
            if (cursor >= payload.size() || payload[cursor] != L'|')
                return false;
            ++cursor;

            MathClipboardFragmentEntry entry;
            entry.relativeStart = (LONG)relativeStart;
            if (!MathObject::ParseString(payload, cursor, entry.payload))
                return false;
            fragment.entries.push_back(std::move(entry));
        }

        if (cursor >= payload.size() || payload[cursor] != L']')
            return false;
        ++cursor;
        return cursor == payload.size();
    }

    static bool TryGetClipboardObjectIndex(HWND hwnd, size_t& outIndex)
    {
        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();
        auto& objects = mgr.GetObjects();

        if (state.active && state.objectIndex < objects.size())
        {
            outIndex = state.objectIndex;
            return true;
        }

        DWORD selStart = 0, selEnd = 0;
        SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);
        if (selStart == selEnd)
        {
            if (mgr.IsPosInsideAnyObject((LONG)selStart, &outIndex))
                return true;
            if (selStart > 0 && mgr.IsPosInsideAnyObject((LONG)selStart - 1, &outIndex))
                return true;
            return false;
        }

        for (size_t index = 0; index < objects.size(); ++index)
        {
            const MathObject& obj = objects[index];
            if ((LONG)selStart == obj.barStart && (LONG)selEnd == obj.barStart + obj.barLen)
            {
                outIndex = index;
                return true;
            }
        }
        return false;
    }

    static bool TryCollectSelectedMathObjects(HWND hwnd, LONG& outSelStart, LONG& outSelEnd, std::vector<size_t>& outIndices)
    {
        auto& objects = MathManager::Get().GetObjects();
        DWORD selStart = 0, selEnd = 0;
        SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);
        outSelStart = (LONG)selStart;
        outSelEnd = (LONG)selEnd;
        outIndices.clear();
        if (selEnd <= selStart)
            return false;

        for (size_t index = 0; index < objects.size(); ++index)
        {
            const MathObject& obj = objects[index];
            const LONG objEnd = obj.barStart + obj.barLen;
            const bool overlaps = !((LONG)selEnd <= obj.barStart || (LONG)selStart >= objEnd);
            if (!overlaps)
                continue;
            if ((LONG)selStart > obj.barStart || (LONG)selEnd < objEnd)
                return false;
            outIndices.push_back(index);
        }
        return !outIndices.empty();
    }

    static std::wstring BuildSelectionPlainTextFallback(HWND hwnd, LONG start, LONG end, const std::vector<size_t>& objectIndices)
    {
        auto& objects = MathManager::Get().GetObjects();
        std::wstring text;
        LONG cursor = start;
        for (size_t index : objectIndices)
        {
            const MathObject& obj = objects[index];
            if (cursor < obj.barStart)
                text += ExtractTextRange(hwnd, cursor, obj.barStart);
            text += obj.BuildPlainTextFallback();
            cursor = obj.barStart + obj.barLen;
        }
        if (cursor < end)
            text += ExtractTextRange(hwnd, cursor, end);
        return text;
    }

    static bool BuildClipboardFragmentFromSelection(HWND hwnd, LONG start, LONG end, const std::vector<size_t>& objectIndices, MathClipboardFragment& outFragment, std::wstring& outPlainText)
    {
        outFragment.rawText = ExtractTextRange(hwnd, start, end);
        outFragment.entries.clear();
        outFragment.entries.reserve(objectIndices.size());

        auto& objects = MathManager::Get().GetObjects();
        for (size_t index : objectIndices)
        {
            const MathObject& obj = objects[index];
            MathClipboardFragmentEntry entry;
            entry.relativeStart = obj.barStart - start;
            entry.payload = obj.SerializeTransferPayload();
            outFragment.entries.push_back(std::move(entry));
        }

        outPlainText = BuildSelectionPlainTextFallback(hwnd, start, end, objectIndices);
        return true;
    }

    static void GetAnchorSpecForObject(const MathObject& obj, std::wstring& outAnchor, LONG& outAnchorLen, bool& outNormalHeight, LONG forcedAnchorLen = 0)
    {
        outNormalHeight = (obj.type == MathType::SquareRoot || obj.type == MathType::AbsoluteValue);
        if (obj.type == MathType::Fraction)
        {
            outAnchorLen = forcedAnchorLen > 0 ? forcedAnchorLen : (LONG)std::max<size_t>(3, std::max(obj.SlotText(1).size(), obj.SlotText(2).size()));
            outAnchor.assign((size_t)outAnchorLen, L'\u2500');
            return;
        }

        outAnchorLen = forcedAnchorLen > 0 ? forcedAnchorLen : 5;
        outAnchor.assign((size_t)outAnchorLen, L'\u00A0');
    }

    static bool OverlayStructuredObjectAtRange(HWND hwnd, LONG start, MathObject obj, LONG forcedAnchorLen = 0)
    {
        auto& objects = MathManager::Get().GetObjects();

        std::wstring anchorText;
        LONG anchorLen = 0;
        bool normalHeight = false;
        GetAnchorSpecForObject(obj, anchorText, anchorLen, normalHeight, forcedAnchorLen);

        SendMessage(hwnd, EM_SETSEL, (WPARAM)start, (LPARAM)(start + anchorLen));

        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        SendMessage(hwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
        const LONG originalFontHeight = cf.yHeight;
        if (!normalHeight)
        {
            cf.dwMask |= CFM_SIZE;
            cf.yHeight *= 2;
            SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
        }

        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)anchorText.c_str());
        if (obj.type == MathType::Fraction)
            HideAnchorChars(hwnd, start, anchorLen);

        PARAFORMAT2 pf2 = {};
        pf2.cbSize = sizeof(pf2);
        pf2.dwMask = PFM_SPACEBEFORE | PFM_SPACEAFTER;
        if (normalHeight)
        {
            pf2.dySpaceBefore = (LONG)(originalFontHeight * 0.5);
            pf2.dySpaceAfter = (LONG)(originalFontHeight * 0.3);
        }
        else
        {
            pf2.dySpaceBefore = (LONG)(originalFontHeight * 1.5);
            pf2.dySpaceAfter = (LONG)(originalFontHeight * 1.5);
        }
        SendMessage(hwnd, EM_SETPARAFORMAT, 0, (LPARAM)&pf2);

        obj.barStart = start;
        obj.barLen = anchorLen;
        objects.push_back(std::move(obj));
        return true;
    }

    static bool InsertStructuredObjectAtSelection(HWND hwnd, MathObject obj)
    {
        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();
        auto& objects = mgr.GetObjects();

        DWORD selStart = 0, selEnd = 0;
        SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

        if (state.active)
        {
            ShowCaret(hwnd);
            state.active = false;
            state.activeNodePath.clear();
        }

        mgr.DeleteObjectsInRange((LONG)selStart, (LONG)selEnd);

        std::wstring anchorText;
        LONG anchorLen = 0;
        bool normalHeight = false;
        GetAnchorSpecForObject(obj, anchorText, anchorLen, normalHeight);

        ScopedNoRedraw noRedraw(hwnd);
        SendMessage(hwnd, EM_SETSEL, (WPARAM)selStart, (LPARAM)selEnd);

        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        SendMessage(hwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
        const LONG originalFontHeight = cf.yHeight;
        if (!normalHeight)
        {
            cf.dwMask |= CFM_SIZE;
            cf.yHeight *= 2;
            SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
        }

        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)anchorText.c_str());
        if (obj.type == MathType::Fraction)
            HideAnchorChars(hwnd, (LONG)selStart, anchorLen);

        PARAFORMAT2 pf2 = {};
        pf2.cbSize = sizeof(pf2);
        pf2.dwMask = PFM_SPACEBEFORE | PFM_SPACEAFTER;
        if (normalHeight)
        {
            pf2.dySpaceBefore = (LONG)(originalFontHeight * 0.5);
            pf2.dySpaceAfter = (LONG)(originalFontHeight * 0.3);
        }
        else
        {
            pf2.dySpaceBefore = (LONG)(originalFontHeight * 1.5);
            pf2.dySpaceAfter = (LONG)(originalFontHeight * 1.5);
        }
        SendMessage(hwnd, EM_SETPARAFORMAT, 0, (LPARAM)&pf2);

        mgr.ShiftObjectsAfter((LONG)selEnd, anchorLen - (LONG)(selEnd - selStart));
        obj.barStart = (LONG)selStart;
        obj.barLen = anchorLen;
        objects.push_back(std::move(obj));

        SendMessage(hwnd, EM_SETSEL, (WPARAM)(selStart + anchorLen), (LPARAM)(selStart + anchorLen));
        RestoreTypingFormat(hwnd);
        RequestMathRepaint(hwnd);
        return true;
    }

    static bool CopyMathObjectToClipboard(HWND hwnd, const MathObject& obj)
    {
        if (!OpenClipboard(hwnd))
            return false;

        const UINT clipboardFormat = GetMathClipboardFormat();
        const std::wstring payload = obj.SerializeTransferPayload();
        const std::wstring plainText = obj.BuildPlainTextFallback();

        EmptyClipboard();
        bool ok = SetClipboardPayload(clipboardFormat, payload);
        ok = SetClipboardUnicodeText(plainText) && ok;
        CloseClipboard();
        return ok;
    }

    static bool CopyMathFragmentToClipboard(HWND hwnd, LONG start, LONG end, const std::vector<size_t>& objectIndices)
    {
        MathClipboardFragment fragment;
        std::wstring plainText;
        if (!BuildClipboardFragmentFromSelection(hwnd, start, end, objectIndices, fragment, plainText))
            return false;

        if (!OpenClipboard(hwnd))
            return false;

        EmptyClipboard();
        bool ok = SetClipboardPayload(GetMathFragmentClipboardFormat(), SerializeClipboardFragment(fragment));
        ok = SetClipboardUnicodeText(plainText) && ok;
        CloseClipboard();
        return ok;
    }

    static bool TryPasteMathObjectFromClipboard(HWND hwnd)
    {
        if (!OpenClipboard(hwnd))
            return false;

        std::wstring payload;
        const bool hasPayload = ReadClipboardPayload(GetMathClipboardFormat(), payload);
        CloseClipboard();
        if (!hasPayload)
            return false;

        MathObject obj;
        if (!MathObject::TryDeserializeTransferPayload(payload, obj))
            return false;
        return InsertStructuredObjectAtSelection(hwnd, std::move(obj));
    }

    static bool TryPasteMathFragmentFromClipboard(HWND hwnd)
    {
        if (!OpenClipboard(hwnd))
            return false;

        std::wstring payload;
        const bool hasPayload = ReadClipboardPayload(GetMathFragmentClipboardFormat(), payload);
        CloseClipboard();
        if (!hasPayload)
            return false;

        MathClipboardFragment fragment;
        if (!TryDeserializeClipboardFragment(payload, fragment))
            return false;

        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();

        DWORD selStart = 0, selEnd = 0;
        SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);
        if (state.active)
        {
            ShowCaret(hwnd);
            state.active = false;
            state.activeNodePath.clear();
        }

        ScopedNoRedraw noRedraw(hwnd);
        mgr.DeleteObjectsInRange((LONG)selStart, (LONG)selEnd);
        SendMessage(hwnd, EM_SETSEL, (WPARAM)selStart, (LPARAM)selEnd);
        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)fragment.rawText.c_str());
        mgr.ShiftObjectsAfter((LONG)selEnd, (LONG)fragment.rawText.size() - (LONG)(selEnd - selStart));

        for (const auto& entry : fragment.entries)
        {
            MathObject obj;
            if (!MathObject::TryDeserializeTransferPayload(entry.payload, obj))
                return false;
            if (!OverlayStructuredObjectAtRange(hwnd, (LONG)selStart + entry.relativeStart, std::move(obj)))
                return false;
        }

        SendMessage(hwnd, EM_SETSEL, (WPARAM)(selStart + fragment.rawText.size()), (LPARAM)(selStart + fragment.rawText.size()));
        RestoreTypingFormat(hwnd);
        RequestMathRepaint(hwnd);
        return true;
    }

    static int GetEditablePartCount(const MathObject& obj)
    {
        switch (obj.type)
        {
        case MathType::Matrix:
        case MathType::Determinant:
            return 4;
        case MathType::Fraction:
            return 2;
        case MathType::SquareRoot:
            return 2;
        case MathType::AbsoluteValue:
            return 1;
        case MathType::Power:
            return 2;
        case MathType::Logarithm:
            return 2;
        case MathType::Sum:
            return 1;
        default:
            return 3;
        }
    }

    static std::wstring& GetActiveEditText(MathObject& obj, int activePart)
    {
        return obj.EditableSlotText(activePart);
    }

    static std::wstring& GetActiveEditText(MathObject& obj, int activePart, const std::vector<size_t>& activeNodePath)
    {
        if (!activeNodePath.empty())
            return obj.EditableLeafText(activePart, &activeNodePath);
        return obj.EditableSlotText(activePart);
    }

    static void RefreshActiveSlotText(MathObject& obj, int activePart)
    {
        obj.RebuildSlotTextFromChildren(MathObject::SlotIndexFromPart(activePart));
    }

    static void SyncFractionAnchorLength(HWND hwnd, MathObject& obj)
    {
        if (obj.type != MathType::Fraction)
            return;

        const LONG requiredLen = (LONG)std::max<size_t>(3, std::max(obj.SlotText(1).size(), obj.SlotText(2).size()));
        if (requiredLen == obj.barLen)
            return;

        const LONG originalLen = obj.barLen;
        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + originalLen));
        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)requiredLen, L'\u2500').c_str());
        HideAnchorChars(hwnd, obj.barStart, requiredLen);
        MathManager::Get().ShiftObjectsAfter(obj.barStart + originalLen, requiredLen - originalLen);
        obj.barLen = requiredLen;
    }

    static bool TryInsertNestedMathCommand(MathObject& obj, MathTypingState& state)
    {
        const struct NestedCommand { const wchar_t* trigger; MathNodeKind kind; size_t initialSlot; } commands[] = {
            { L"\\sqrt", MathNodeKind::SquareRoot, 0 },
            { L"\\frac", MathNodeKind::Fraction, 0 },
            { L"\\pow",  MathNodeKind::Power, 0 },
            { L"\\abs",  MathNodeKind::AbsoluteValue, 0 },
            { L"\\log",  MathNodeKind::Logarithm, 1 }
        };

        for (const auto& command : commands)
        {
            std::vector<size_t> nextPath;
            if (obj.InsertNestedNode(state.activePart, state.activeNodePath, command.trigger, command.kind, nextPath, command.initialSlot))
            {
                state.activeNodePath = std::move(nextPath);
                return true;
            }
        }
        return false;
    }

    static bool BuildMathDocumentSnapshot(HWND hwnd, MathDocumentSnapshot& outSnapshot)
    {
        const LONG textLen = GetWindowTextLengthW(hwnd);
        outSnapshot.rawText = ExtractTextRange(hwnd, 0, textLen);
        outSnapshot.entries.clear();

        auto& objects = MathManager::Get().GetObjects();
        outSnapshot.entries.reserve(objects.size());
        for (const auto& obj : objects)
        {
            MathDocumentEntry entry;
            entry.start = obj.barStart;
            entry.length = obj.barLen;
            entry.payload = obj.SerializeTransferPayload();
            outSnapshot.entries.push_back(std::move(entry));
        }
        return true;
    }

    static bool RestoreMathDocumentSnapshot(HWND hwnd, const MathDocumentSnapshot& snapshot)
    {
        if (!hwnd)
            return false;

        SetWindowTextW(hwnd, snapshot.rawText.c_str());
        ResetMathSupport();
        SendMessage(hwnd, EM_SETSEL, 0, 0);

        for (const auto& entry : snapshot.entries)
        {
            MathObject obj;
            if (!MathObject::TryDeserializeTransferPayload(entry.payload, obj))
                return false;
            if (!OverlayStructuredObjectAtRange(hwnd, entry.start, std::move(obj), entry.length))
                return false;
        }

        SendMessage(hwnd, EM_SETSEL, 0, 0);
        RequestMathRepaint(hwnd);
        return true;
    }

    static void SendChars(HWND hwnd, const std::wstring& text)
    {
        for (wchar_t ch : text)
            SendMessage(hwnd, WM_CHAR, (WPARAM)ch, 0);
    }

    static bool TryInsertFunctionTemplate(HWND hwnd, LONG selEndBefore, const std::wstring& command)
    {
        const struct TemplateCommand { const wchar_t* name; const wchar_t* replacement; int caretOffset; } templates[] = {
            { L"\\sin",  L"sin()", 4 },
            { L"\\cos",  L"cos()", 4 },
            { L"\\tan",  L"tan()", 4 },
            { L"\\asin", L"asin()", 5 },
            { L"\\acos", L"acos()", 5 },
            { L"\\atan", L"atan()", 5 },
            { L"\\ln",   L"ln()", 3 },
            { L"\\exp",  L"exp()", 4 }
        };

        for (const auto& item : templates)
        {
            if (command != item.name) continue;
            const LONG cmdLen = (LONG)command.size();
            const LONG cmdStart = selEndBefore - cmdLen;
            if (cmdStart < 0) return false;

            ScopedNoRedraw noRedraw(hwnd);
            SendMessage(hwnd, EM_SETSEL, (WPARAM)cmdStart, (LPARAM)selEndBefore);
            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)item.replacement);
            const LONG replacementLen = (LONG)wcslen(item.replacement);
            MathManager::Get().ShiftObjectsAfter(selEndBefore, replacementLen - cmdLen);
            SendMessage(hwnd, EM_SETSEL, (WPARAM)(cmdStart + item.caretOffset), (LPARAM)(cmdStart + item.caretOffset));
            RequestMathRepaint(hwnd);
            return true;
        }

        return false;
    }

    static void UpdateResultIfPresent(HWND hwnd, size_t objIdx)
    {
        auto& mgr = MathManager::Get();
        auto& objects = mgr.GetObjects();
        if (objIdx >= objects.size()) return;
        auto& obj = objects[objIdx];
        if (obj.resultText.empty()) return;

        if (obj.type == MathType::SystemOfEquations) {
            // For system of equations, use the dedicated calculation method
            std::wstring systemResult = mgr.CalculateSystemResult(obj);
            obj.resultText = systemResult; // CalculateSystemResult already includes the equals sign
        } else {
            if (!mgr.CanCalculateResult(obj)) return;
            obj.resultText = mgr.CalculateFormattedResult(obj);
        }
        RequestMathRepaint(hwnd);
    }

    static void TriggerCalculation(HWND hwnd, size_t objIdx)
    {
        auto& mgr = MathManager::Get();
        auto& objects = mgr.GetObjects();
        if (objIdx >= objects.size()) return;
        auto& obj = objects[objIdx];

        if (obj.type == MathType::SystemOfEquations) {
            // For system of equations, use the dedicated calculation method
            std::wstring systemResult = MathManager::Get().CalculateSystemResult(obj);
            obj.resultText = systemResult; // CalculateSystemResult already includes the equals sign
            SendMessage(hwnd, EM_SETSEL, obj.barStart + obj.barLen, obj.barStart + obj.barLen);
            RestoreTypingFormat(hwnd);
            RequestMathRepaint(hwnd);
            return;
        }

        if (!mgr.CanCalculateResult(obj)) return;

        obj.resultText = mgr.CalculateFormattedResult(obj);
        SendMessage(hwnd, EM_SETSEL, obj.barStart + obj.barLen, obj.barStart + obj.barLen);
        RestoreTypingFormat(hwnd);
        RequestMathRepaint(hwnd);
    }

    static void ClearUnitSuggestionPopup()
    {
        g_unitSuggestionPopup = {};
    }

    static COLORREF GetEditorBackgroundColor(HWND hwnd)
    {
        COLORREF bk = (COLORREF)SendMessage(hwnd, EM_SETBKGNDCOLOR, 0, 0);
        SendMessage(hwnd, EM_SETBKGNDCOLOR, 0, (LPARAM)bk);
        return bk;
    }

    static bool EqualsIgnoreCase(const std::wstring& left, const std::wstring& right)
    {
        if (left.size() != right.size())
            return false;

        for (size_t index = 0; index < left.size(); ++index)
        {
            if (towlower(left[index]) != towlower(right[index]))
                return false;
        }
        return true;
    }

    static void HideUnitSuggestionPopup(HWND hwnd)
    {
        if (!g_unitSuggestionPopup.visible)
            return;
        ClearUnitSuggestionPopup();
        InvalidateMathOverlay(hwnd);
    }

    static bool TryBuildUnitSuggestionContext(const std::wstring& text, bool forceAll, UnitSuggestionContext& outContext)
    {
        outContext = {};

        const size_t textLen = text.size();
        if (textLen == 0)
        {
            if (!forceAll)
                return false;
            outContext.showAll = true;
            return true;
        }

        size_t trimmedEnd = textLen;
        while (trimmedEnd > 0 && iswspace(text[trimmedEnd - 1]))
            --trimmedEnd;
        const bool endsWithSpace = trimmedEnd < textLen;

        size_t prefixStart = trimmedEnd;
        while (prefixStart > 0 && iswalpha(text[prefixStart - 1]))
            --prefixStart;

        if (prefixStart < trimmedEnd)
        {
            if (prefixStart > 0 && text[prefixStart - 1] == L'\\')
                return false;
            outContext.replaceStart = prefixStart;
            outContext.prefix = text.substr(prefixStart, trimmedEnd - prefixStart);
            return true;
        }

        if (!forceAll && !endsWithSpace)
            return false;

        if (!forceAll)
        {
            if (trimmedEnd == 0)
                return false;

            const wchar_t marker = text[trimmedEnd - 1];
            if (!(iswdigit(marker) || marker == L')' || marker == L'}' || marker == L'*' || marker == L'/'))
                return false;
        }

        outContext.showAll = true;
        outContext.replaceStart = textLen;
        return true;
    }

    static void UpdateUnitSuggestionLayout(HWND hwnd)
    {
        if (!g_unitSuggestionPopup.visible)
            return;

        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();
        auto& objects = mgr.GetObjects();
        if (!state.active || state.objectIndex >= objects.size() || g_unitSuggestionPopup.items.empty())
        {
            ClearUnitSuggestionPopup();
            return;
        }

        HDC hdc = GetDC(hwnd);
        if (!hdc)
            return;

        HFONT baseFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);
        if (!baseFont)
            baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT oldFont = (HFONT)SelectObject(hdc, baseFont);

        TEXTMETRICW tm = {};
        GetTextMetricsW(hdc, &tm);
        const int popupEdgePadding = 10;
        const int popupGap = 8;
        const int caretLead = 6;

        int width = 0;
        for (const auto& item : g_unitSuggestionPopup.items)
        {
            SIZE textSize = {};
            GetTextExtentPoint32W(hdc, item.c_str(), (int)item.size(), &textSize);
            width = (std::max)(width, (int)textSize.cx);
        }

        const int itemHeight = tm.tmHeight + 6;
        const size_t maxVisibleItems = 8;
        const size_t visibleCount = (std::min)(g_unitSuggestionPopup.items.size(), maxVisibleItems);
        if (visibleCount == 0)
        {
            SelectObject(hdc, oldFont);
            ReleaseDC(hwnd, hdc);
            ClearUnitSuggestionPopup();
            return;
        }

        if (g_unitSuggestionPopup.selectedIndex < g_unitSuggestionPopup.topIndex)
            g_unitSuggestionPopup.topIndex = g_unitSuggestionPopup.selectedIndex;
        if (g_unitSuggestionPopup.selectedIndex >= g_unitSuggestionPopup.topIndex + visibleCount)
            g_unitSuggestionPopup.topIndex = g_unitSuggestionPopup.selectedIndex - visibleCount + 1;

        RECT objectRect = {};
        const bool hasObjectRect = MathRenderer::TryGetObjectBounds(hwnd, hdc, objects[state.objectIndex], objectRect);

        RECT anchorRect = { popupEdgePadding, popupEdgePadding, popupEdgePadding + width, popupEdgePadding + tm.tmHeight };
        POINT caretPt = {};
        const bool hasCaretPt = MathRenderer::TryGetActiveCaretPoint(hwnd, hdc, objects[state.objectIndex], state.objectIndex, state, caretPt);
        if (hasCaretPt)
        {
            anchorRect.left = caretPt.x + caretLead;
            anchorRect.top = caretPt.y - tm.tmAscent;
            anchorRect.right = anchorRect.left + width;
            anchorRect.bottom = caretPt.y + tm.tmDescent;
            if (hasObjectRect)
            {
                anchorRect.top = (std::min)(anchorRect.top, objectRect.top);
                anchorRect.bottom = (std::max)(anchorRect.bottom, objectRect.bottom);
            }
        }
        else if (hasObjectRect)
        {
            anchorRect = objectRect;
        }
        else
        {
            POINT anchorPt = {};
            if (MathRenderer::TryGetCharPos(hwnd, objects[state.objectIndex].barStart, anchorPt))
            {
                anchorRect.left = anchorPt.x + caretLead;
                anchorRect.top = anchorPt.y;
                anchorRect.right = anchorRect.left + width;
                anchorRect.bottom = anchorPt.y + tm.tmHeight;
            }
        }

        RECT client = {};
        GetClientRect(hwnd, &client);

        const int availableWidth = (std::max)(80, (int)(client.right - client.left - popupEdgePadding * 2));
        width = (std::min)(width + 16, availableWidth);
        const int height = (int)visibleCount * itemHeight + 2;

        int x = anchorRect.left;
        int y = anchorRect.bottom + popupGap;
        if (y + height > client.bottom - popupEdgePadding)
            y = (std::max)((int)client.top + popupEdgePadding, (int)anchorRect.top - height - popupGap);
        if (x + width > client.right - popupEdgePadding)
            x = (std::max)((int)client.left + popupEdgePadding, (int)client.right - width - popupEdgePadding);
        if (x < client.left + popupEdgePadding)
            x = client.left + popupEdgePadding;
        if (y < client.top + popupEdgePadding)
            y = client.top + popupEdgePadding;

        g_unitSuggestionPopup.popupRect = { x, y, x + width, y + height };
        g_unitSuggestionPopup.itemRects.clear();
        g_unitSuggestionPopup.itemRects.reserve(visibleCount);
        for (size_t visibleIndex = 0; visibleIndex < visibleCount; ++visibleIndex)
        {
            RECT itemRect = {
                x + 1,
                y + 1 + (int)visibleIndex * itemHeight,
                x + width - 1,
                y + 1 + (int)(visibleIndex + 1) * itemHeight
            };
            g_unitSuggestionPopup.itemRects.push_back(itemRect);
        }

        SelectObject(hdc, oldFont);
        ReleaseDC(hwnd, hdc);
    }

    static void RefreshUnitSuggestionPopup(HWND hwnd, bool forceAll = false)
    {
        UnitSuggestionPopupState previousState = g_unitSuggestionPopup;

        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();
        auto& objects = mgr.GetObjects();
        if (!state.active || state.objectIndex >= objects.size())
        {
            HideUnitSuggestionPopup(hwnd);
            return;
        }

        MathObject& obj = objects[state.objectIndex];
        const std::wstring& target = GetActiveEditText(obj, state.activePart, state.activeNodePath);

        UnitSuggestionContext context;
        if (!TryBuildUnitSuggestionContext(target, forceAll, context))
        {
            HideUnitSuggestionPopup(hwnd);
            return;
        }

        std::vector<std::wstring> matches = context.showAll ? GetKnownUnitSymbols() : FindMatchingUnitSymbols(context.prefix);
        if (!forceAll && !context.prefix.empty() && matches.size() == 1 && EqualsIgnoreCase(matches[0], context.prefix))
            matches.clear();
        if (matches.empty())
        {
            HideUnitSuggestionPopup(hwnd);
            return;
        }

        size_t selectedIndex = 0;
        if (previousState.visible && previousState.selectedIndex < previousState.items.size())
        {
            const std::wstring& previouslySelected = previousState.items[previousState.selectedIndex];
            auto it = std::find(matches.begin(), matches.end(), previouslySelected);
            if (it != matches.end())
                selectedIndex = (size_t)std::distance(matches.begin(), it);
        }

        g_unitSuggestionPopup.visible = true;
        g_unitSuggestionPopup.replaceStart = context.replaceStart;
        g_unitSuggestionPopup.prefix = context.prefix;
        g_unitSuggestionPopup.items = std::move(matches);
        g_unitSuggestionPopup.selectedIndex = (std::min)(selectedIndex, g_unitSuggestionPopup.items.size() - 1);
        if (!previousState.visible)
            g_unitSuggestionPopup.topIndex = 0;

        UpdateUnitSuggestionLayout(hwnd);

        const bool changed = !previousState.visible ||
            previousState.items != g_unitSuggestionPopup.items ||
            previousState.selectedIndex != g_unitSuggestionPopup.selectedIndex ||
            previousState.topIndex != g_unitSuggestionPopup.topIndex ||
            previousState.replaceStart != g_unitSuggestionPopup.replaceStart ||
            previousState.prefix != g_unitSuggestionPopup.prefix;
        if (changed)
            InvalidateMathOverlay(hwnd);
    }

    static bool MoveUnitSuggestionSelection(HWND hwnd, int delta)
    {
        if (!g_unitSuggestionPopup.visible || g_unitSuggestionPopup.items.empty())
            return false;

        const size_t itemCount = g_unitSuggestionPopup.items.size();
        const size_t originalIndex = g_unitSuggestionPopup.selectedIndex;
        if (delta < 0)
        {
            const size_t magnitude = (size_t)(-delta);
            g_unitSuggestionPopup.selectedIndex = (g_unitSuggestionPopup.selectedIndex > magnitude)
                ? (g_unitSuggestionPopup.selectedIndex - magnitude)
                : 0;
        }
        else if (delta > 0)
        {
            g_unitSuggestionPopup.selectedIndex = (std::min)(itemCount - 1, g_unitSuggestionPopup.selectedIndex + (size_t)delta);
        }

        if (g_unitSuggestionPopup.selectedIndex == originalIndex)
            return true;

        UpdateUnitSuggestionLayout(hwnd);
        InvalidateMathOverlay(hwnd);
        return true;
    }

    static bool ApplySelectedUnitSuggestion(HWND hwnd)
    {
        if (!g_unitSuggestionPopup.visible || g_unitSuggestionPopup.items.empty())
            return false;

        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();
        auto& objects = mgr.GetObjects();
        if (!state.active || state.objectIndex >= objects.size())
        {
            HideUnitSuggestionPopup(hwnd);
            return false;
        }

        MathObject& obj = objects[state.objectIndex];
        std::wstring& target = GetActiveEditText(obj, state.activePart, state.activeNodePath);
        if (g_unitSuggestionPopup.replaceStart > target.size())
            return false;

        target.erase(g_unitSuggestionPopup.replaceStart);
        target += g_unitSuggestionPopup.items[g_unitSuggestionPopup.selectedIndex];
        RefreshActiveSlotText(obj, state.activePart);
        SyncFractionAnchorLength(hwnd, obj);

        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
        UpdateResultIfPresent(hwnd, state.objectIndex);
        ClearUnitSuggestionPopup();
        InvalidateMathOverlay(hwnd);
        return true;
    }

    static bool TryHandleUnitSuggestionClick(HWND hwnd, POINT pt)
    {
        if (!g_unitSuggestionPopup.visible)
            return false;

        if (!PtInRect(&g_unitSuggestionPopup.popupRect, pt))
        {
            HideUnitSuggestionPopup(hwnd);
            return false;
        }

        for (size_t visibleIndex = 0; visibleIndex < g_unitSuggestionPopup.itemRects.size(); ++visibleIndex)
        {
            if (!PtInRect(&g_unitSuggestionPopup.itemRects[visibleIndex], pt))
                continue;
            g_unitSuggestionPopup.selectedIndex = g_unitSuggestionPopup.topIndex + visibleIndex;
            return ApplySelectedUnitSuggestion(hwnd);
        }

        return true;
    }

    static void DrawUnitSuggestionPopup(HWND hwnd, HDC hdc)
    {
        if (!g_unitSuggestionPopup.visible || g_unitSuggestionPopup.items.empty())
            return;

        UpdateUnitSuggestionLayout(hwnd);
        if (!g_unitSuggestionPopup.visible || g_unitSuggestionPopup.items.empty())
            return;

        const COLORREF editorBg = GetEditorBackgroundColor(hwnd);
        const int brightness = (GetRValue(editorBg) * 299 + GetGValue(editorBg) * 587 + GetBValue(editorBg) * 114) / 1000;
        const bool darkBackground = brightness < 128;
        const COLORREF popupBg = darkBackground ? RGB(34, 34, 34) : RGB(252, 252, 252);
        const COLORREF borderColor = darkBackground ? RGB(88, 88, 88) : RGB(188, 188, 188);
        const COLORREF selectedBg = darkBackground ? RGB(62, 83, 112) : RGB(221, 236, 255);
        const COLORREF normalText = MathRenderer::GetDefaultTextColor(hwnd);
        const COLORREF selectedText = MathRenderer::GetActiveColor(hwnd);

        HBRUSH popupBrush = CreateSolidBrush(popupBg);
        FillRect(hdc, &g_unitSuggestionPopup.popupRect, popupBrush);
        DeleteObject(popupBrush);

        HBRUSH borderBrush = CreateSolidBrush(borderColor);
        FrameRect(hdc, &g_unitSuggestionPopup.popupRect, borderBrush);
        DeleteObject(borderBrush);

        HFONT baseFont = (HFONT)SendMessage(hwnd, WM_GETFONT, 0, 0);
        if (!baseFont)
            baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);
        HFONT oldFont = (HFONT)SelectObject(hdc, baseFont);

        SetBkMode(hdc, TRANSPARENT);
        for (size_t visibleIndex = 0; visibleIndex < g_unitSuggestionPopup.itemRects.size(); ++visibleIndex)
        {
            const size_t itemIndex = g_unitSuggestionPopup.topIndex + visibleIndex;
            RECT itemRect = g_unitSuggestionPopup.itemRects[visibleIndex];
            if (itemIndex == g_unitSuggestionPopup.selectedIndex)
            {
                HBRUSH selectedBrush = CreateSolidBrush(selectedBg);
                FillRect(hdc, &itemRect, selectedBrush);
                DeleteObject(selectedBrush);
                SetTextColor(hdc, selectedText);
            }
            else
            {
                SetTextColor(hdc, normalText);
            }

            TextOutW(hdc, itemRect.left + 6, itemRect.top + 3, g_unitSuggestionPopup.items[itemIndex].c_str(), (int)g_unitSuggestionPopup.items[itemIndex].size());
        }

        SelectObject(hdc, oldFont);
    }

    static LRESULT CALLBACK MathRichEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        auto& mgr = MathManager::Get();
        auto& state = mgr.GetState();
        auto& objects = mgr.GetObjects();

        switch (uMsg)
        {
        case WM_SETCURSOR:
        {
            if (LOWORD(lParam) == HTCLIENT)
            {
                POINT pt; GetCursorPos(&pt); ScreenToClient(hwnd, &pt);
                if (g_unitSuggestionPopup.visible && PtInRect(&g_unitSuggestionPopup.popupRect, pt))
                {
                    SetCursor(LoadCursor(nullptr, IDC_ARROW));
                    return TRUE;
                }
                HDC hdc = GetDC(hwnd); size_t idx = 0; int part = 0;
                bool hit = MathRenderer::GetHitPart(hwnd, hdc, pt, &idx, &part);
                if (!hit) {
                    POINTL ptl = { pt.x, pt.y };
                    LRESULT charIdx = SendMessage(hwnd, EM_CHARFROMPOS, 0, (LPARAM)&ptl);
                    hit = mgr.IsPosInsideAnyObject((LONG)charIdx);
                }
                ReleaseDC(hwnd, hdc);
                if (hit) { SetCursor(LoadCursor(nullptr, IDC_HAND)); return TRUE; }
            }
            break;
        }

        case WM_SETFOCUS:
        {
            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            if (state.active) HideCaret(hwnd);
            return res;
        }

        case WM_KILLFOCUS:
        {
            HideUnitSuggestionPopup(hwnd);
            return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
        }

        case WM_LBUTTONDOWN:
        {
            POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
            if (TryHandleUnitSuggestionClick(hwnd, pt))
                return 0;

            HDC hdc = GetDC(hwnd); size_t idx = 0; int part = 0; std::vector<size_t> nodePath;
            bool hit = MathRenderer::GetHitPart(hwnd, hdc, pt, &idx, &part, &nodePath);
            ReleaseDC(hwnd, hdc);

            g_currentNumber.clear(); 
            g_currentCommand.clear();

            if (hit)
            {
                SetFocus(hwnd); 
                if (!state.active) HideCaret(hwnd);
                state.active = true;
                state.objectIndex = idx;
                state.activePart = part;
                state.activeNodePath = std::move(nodePath);
                SendMessage(hwnd, EM_SETSEL, (WPARAM)objects[idx].barStart, (LPARAM)objects[idx].barStart);
                RefreshUnitSuggestionPopup(hwnd);
                RequestMathRepaint(hwnd);
                return 0; 
            }
            else
            {
                HideUnitSuggestionPopup(hwnd);
                if (state.active) { ShowCaret(hwnd); state.active = false; state.activeNodePath.clear(); RequestMathRepaint(hwnd); }
            }
            break; 
        }

        case WM_PAINT:
        {
            if (objects.empty())
            {
                // No math objects — pass through to RichEdit directly.
                return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            }

            // Double-buffer: compose RichEdit content + math overlays
            // off-screen, then blit once to prevent flicker.
            PAINTSTRUCT ps;
            HDC hdcPaint = BeginPaint(hwnd, &ps);

            RECT rc;
            GetClientRect(hwnd, &rc);

            HDC hdcMem = CreateCompatibleDC(hdcPaint);
            HBITMAP hbm = CreateCompatibleBitmap(hdcPaint, rc.right, rc.bottom);
            HBITMAP hbmOld = (HBITMAP)SelectObject(hdcMem, hbm);

            // Render RichEdit content into back-buffer.
            CallWindowProc(g_originalProc, hwnd, WM_PRINTCLIENT, (WPARAM)hdcMem, PRF_CLIENT);

            // Draw math overlays on top.
            if (state.active) HideCaret(hwnd);
            for (size_t i = 0; i < objects.size(); ++i)
                MathRenderer::Draw(hwnd, hdcMem, objects[i], i, state);
            DrawUnitSuggestionPopup(hwnd, hdcMem);

            // Single
            BitBlt(hdcPaint, 0, 0, rc.right, rc.bottom, hdcMem, 0, 0, SRCCOPY);

            SelectObject(hdcMem, hbmOld);
            DeleteObject(hbm);
            DeleteDC(hdcMem);

            EndPaint(hwnd, &ps);
            return 0;
        }

        case WM_PRINTCLIENT:
        {
            const LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            HDC hdc = (HDC)wParam;
            if (hdc)
            {
                if (state.active) HideCaret(hwnd);
                for (size_t i = 0; i < objects.size(); ++i)
                    MathRenderer::Draw(hwnd, hdc, objects[i], i, state);
                DrawUnitSuggestionPopup(hwnd, hdc);
            }
            return res;
        }

        case WM_MOUSEWHEEL:
        case WM_VSCROLL:
        case WM_HSCROLL:
        {
            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            if (g_unitSuggestionPopup.visible)
                InvalidateMathOverlay(hwnd);
            else
                RequestMathRepaint(hwnd);
            return res;
        }

        case WM_COPY:
        case WM_CUT:
        {
            LONG selStart = 0;
            LONG selEnd = 0;
            std::vector<size_t> fragmentIndices;
            if (TryCollectSelectedMathObjects(hwnd, selStart, selEnd, fragmentIndices) && CopyMathFragmentToClipboard(hwnd, selStart, selEnd, fragmentIndices))
            {
                if (uMsg == WM_CUT)
                {
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)selStart, (LPARAM)selEnd);
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                    mgr.DeleteObjectsInRange(selStart, selEnd);
                    mgr.ShiftObjectsAfter(selEnd, selStart - selEnd);
                    state.active = false;
                    state.activeNodePath.clear();
                    ShowCaret(hwnd);
                    RequestMathRepaint(hwnd);
                }
                return 0;
            }

            size_t objectIndex = 0;
            if (TryGetClipboardObjectIndex(hwnd, objectIndex) && objectIndex < objects.size() && CopyMathObjectToClipboard(hwnd, objects[objectIndex]))
            {
                if (uMsg == WM_CUT)
                {
                    const MathObject removedObject = objects[objectIndex];
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)removedObject.barStart, (LPARAM)(removedObject.barStart + removedObject.barLen));
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                    mgr.ShiftObjectsAfter(removedObject.barStart + 1, -removedObject.barLen);
                    objects.erase(objects.begin() + objectIndex);
                    state.active = false;
                    state.activeNodePath.clear();
                    ShowCaret(hwnd);
                    RequestMathRepaint(hwnd);
                }
                return 0;
            }
            break;
        }

        case WM_PASTE:
            if (TryPasteMathFragmentFromClipboard(hwnd))
                return 0;
            if (TryPasteMathObjectFromClipboard(hwnd))
                return 0;
            break;

        case WM_KEYDOWN:
        {
            const LONG lenBefore = (LONG)GetWindowTextLengthW(hwnd);
            DWORD selStart = 0, selEnd = 0;
            SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

            if (state.active && (GetKeyState(VK_CONTROL) & 0x8000) != 0 && wParam == VK_SPACE)
            {
                RefreshUnitSuggestionPopup(hwnd, true);
                return 0;
            }

            if (g_unitSuggestionPopup.visible)
            {
                if (wParam == VK_UP)
                    return MoveUnitSuggestionSelection(hwnd, -1) ? 0 : 0;
                if (wParam == VK_DOWN)
                    return MoveUnitSuggestionSelection(hwnd, 1) ? 0 : 0;
                if (wParam == VK_PRIOR)
                    return MoveUnitSuggestionSelection(hwnd, -5) ? 0 : 0;
                if (wParam == VK_NEXT)
                    return MoveUnitSuggestionSelection(hwnd, 5) ? 0 : 0;
                if (wParam == VK_ESCAPE)
                {
                    HideUnitSuggestionPopup(hwnd);
                    return 0;
                }
                if (wParam == VK_RETURN)
                {
                    if (ApplySelectedUnitSuggestion(hwnd))
                        return 0;
                }
            }

            if ((GetKeyState(VK_CONTROL) & 0x8000) != 0)
            {
                if (wParam == 'C' || wParam == 'X')
                {
                    LONG selStart = 0;
                    LONG selEnd = 0;
                    std::vector<size_t> fragmentIndices;
                    if (TryCollectSelectedMathObjects(hwnd, selStart, selEnd, fragmentIndices) && CopyMathFragmentToClipboard(hwnd, selStart, selEnd, fragmentIndices))
                    {
                        if (wParam == 'X')
                        {
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)selStart, (LPARAM)selEnd);
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                            mgr.DeleteObjectsInRange(selStart, selEnd);
                            mgr.ShiftObjectsAfter(selEnd, selStart - selEnd);
                            state.active = false;
                            state.activeNodePath.clear();
                            ShowCaret(hwnd);
                            RequestMathRepaint(hwnd);
                        }
                        return 0;
                    }

                    size_t objectIndex = 0;
                    if (TryGetClipboardObjectIndex(hwnd, objectIndex) && objectIndex < objects.size() && CopyMathObjectToClipboard(hwnd, objects[objectIndex]))
                    {
                        if (wParam == 'X')
                        {
                            const MathObject removedObject = objects[objectIndex];
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)removedObject.barStart, (LPARAM)(removedObject.barStart + removedObject.barLen));
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                            mgr.ShiftObjectsAfter(removedObject.barStart + 1, -removedObject.barLen);
                            objects.erase(objects.begin() + objectIndex);
                            state.active = false;
                            state.activeNodePath.clear();
                            ShowCaret(hwnd);
                            RequestMathRepaint(hwnd);
                        }
                        return 0;
                    }
                }
                else if (wParam == 'V')
                {
                    if (TryPasteMathFragmentFromClipboard(hwnd))
                        return 0;
                    if (TryPasteMathObjectFromClipboard(hwnd))
                        return 0;
                }
            }

            if (wParam == VK_BACK || wParam == VK_DELETE) {
                if (selStart != selEnd) mgr.DeleteObjectsInRange((LONG)selStart, (LONG)selEnd);
            }

            if (wParam == VK_RETURN)
            {
                if (state.active) {
                    auto& obj = objects[state.objectIndex];
                    LONG afterObj = obj.barStart + obj.barLen;
                    HideUnitSuggestionPopup(hwnd);
                    state.active = false; 
                    state.activeNodePath.clear();
                    

                    // For system of equations, trigger calculation on Enter regardless of position
                    if (obj.type == MathType::SystemOfEquations) {
                        // Simple validation - just make sure both parts exist
                        if (obj.SlotText(1).empty() || obj.SlotText(2).empty()) {
                            return 0;
                        }

                        // Proceed with calculation
                        TriggerCalculation(hwnd, state.objectIndex);
                        RequestMathRepaint(hwnd);  // Force immediate UI refresh
                    }
                    UpdateResultIfPresent(hwnd, state.objectIndex);
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)afterObj, (LPARAM)afterObj);
                    RestoreTypingFormat(hwnd);
                    ShowCaret(hwnd); RequestMathRepaint(hwnd);
                    g_suppressNextChar = true;
                    return 0; 
                }
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                const LONG delta = (LONG)GetWindowTextLengthW(hwnd) - lenBefore;
                const bool affectsObjects = (delta != 0) && HasObjectAtOrAfter(objects, (LONG)selEnd);
                mgr.ShiftObjectsAfter((LONG)selEnd, delta);
                if (affectsObjects) RequestMathRepaint(hwnd);
                return res;
            }

            if (wParam == VK_BACK)
            {
                if (state.active) return 0;
                size_t objIdx = 0;
                if (mgr.IsPosInsideAnyObject((LONG)selEnd - 1, &objIdx)) {
                    auto& obj = objects[objIdx];
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + obj.barLen));
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                    mgr.ShiftObjectsAfter(obj.barStart + 1, -obj.barLen);
                    objects.erase(objects.begin() + objIdx);
                    RequestMathRepaint(hwnd);
                    return 0;
                }
            }

            if (wParam == VK_LEFT || wParam == VK_RIGHT || wParam == VK_UP || wParam == VK_DOWN ||
                wParam == VK_HOME || wParam == VK_END || wParam == VK_PRIOR || wParam == VK_NEXT || wParam == VK_TAB)
            {
                g_currentNumber.clear(); g_currentCommand.clear();
                if (state.active) {
                    auto& obj = objects[state.objectIndex];
                    if (wParam == VK_TAB) {
                        const bool reverse = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
                        if (!state.activeNodePath.empty() && obj.MoveToSiblingSlot(state.activePart, state.activeNodePath, reverse ? -1 : 1)) {
                            HideUnitSuggestionPopup(hwnd);
                            RequestMathRepaint(hwnd);
                            return 0;
                        }
                        state.activeNodePath.clear();
                        int maxP = GetEditablePartCount(objects[state.objectIndex]);
                        if (reverse)
                            state.activePart = (state.activePart + maxP - 2) % maxP + 1;
                        else
                            state.activePart = (state.activePart % maxP) + 1;
                        if (obj.EnterFirstStructuredLeaf(state.activePart, state.activeNodePath) && !reverse) {
                            HideUnitSuggestionPopup(hwnd);
                            RequestMathRepaint(hwnd);
                            return 0;
                        }
                        HideUnitSuggestionPopup(hwnd);
                        RequestMathRepaint(hwnd);
                        return 0;
                    }
                    if (wParam == VK_RIGHT) {
                        if (!state.activeNodePath.empty()) {
                            if (obj.MoveToSiblingSlot(state.activePart, state.activeNodePath, 1)) {
                                HideUnitSuggestionPopup(hwnd);
                                RequestMathRepaint(hwnd);
                                return 0;
                            }
                            state.activeNodePath.clear();
                            HideUnitSuggestionPopup(hwnd);
                            RequestMathRepaint(hwnd);
                            return 0;
                        }
                        if (obj.EnterFirstStructuredLeaf(state.activePart, state.activeNodePath)) {
                            HideUnitSuggestionPopup(hwnd);
                            RequestMathRepaint(hwnd);
                            return 0;
                        }
                    }
                    if ((wParam == VK_LEFT || wParam == VK_HOME) && !state.activeNodePath.empty()) {
                        state.activeNodePath.clear();
                        HideUnitSuggestionPopup(hwnd);
                        RequestMathRepaint(hwnd);
                        return 0;
                    }
                    if (wParam == VK_END) {
                        HideUnitSuggestionPopup(hwnd);
                        RequestMathRepaint(hwnd);
                        return 0;
                    }
                    LONG afterObj = obj.barStart + obj.barLen;
                    HideUnitSuggestionPopup(hwnd);
                    ShowCaret(hwnd); state.active = false;
                    state.activeNodePath.clear();
                    UpdateResultIfPresent(hwnd, state.objectIndex);
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)afterObj, (LPARAM)afterObj);
                    RestoreTypingFormat(hwnd);
                    RequestMathRepaint(hwnd);
                }
            }
            {
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                const LONG delta = (LONG)GetWindowTextLengthW(hwnd) - lenBefore;
                const bool affectsObjects = (delta != 0) && HasObjectAtOrAfter(objects, (LONG)selEnd);
                mgr.ShiftObjectsAfter((LONG)selEnd, delta);
                if (affectsObjects) RequestMathRepaint(hwnd);
                return res;
            }
        }

        case WM_CHAR:
        {
            const wchar_t ch = (wchar_t)wParam;

            // Suppress the WM_CHAR that follows a WM_KEYDOWN which already
            // handled the key (e.g. Enter exiting active math mode).
            if (g_suppressNextChar) {
                g_suppressNextChar = false;
                return 0;
            }

            const LONG lenBefore = (LONG)GetWindowTextLengthW(hwnd);
            DWORD selStartBefore = 0, selEndBefore = 0;
            SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStartBefore, (LPARAM)&selEndBefore);

            size_t hitIndex = 0;
            if (ch == L'=' && mgr.IsPosInsideAnyObject((LONG)selEndBefore - 1, &hitIndex)) {
                TriggerCalculation(hwnd, hitIndex); return 0;
            }

            if (ch == 0x08 && state.active)
            {
                if (state.objectIndex < objects.size()) {
                    auto& obj = objects[state.objectIndex];
                    std::wstring& target = GetActiveEditText(obj, state.activePart, state.activeNodePath);
                    if (!target.empty()) target.pop_back();
                    else if (!state.activeNodePath.empty()) state.activeNodePath.clear();
                    RefreshActiveSlotText(obj, state.activePart);
                    SyncFractionAnchorLength(hwnd, obj);

                    if (obj.SlotText(1).empty() && obj.SlotText(2).empty() && obj.type == MathType::Fraction) {
                        HideUnitSuggestionPopup(hwnd);
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + obj.barLen));
                        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                        mgr.ShiftObjectsAfter(obj.barStart + 1, -obj.barLen);
                        objects.erase(objects.begin() + state.objectIndex);
                        ShowCaret(hwnd); state.active = false;
                    } else {
                        RefreshUnitSuggestionPopup(hwnd);
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                    }
                    UpdateResultIfPresent(hwnd, state.objectIndex);
                    RequestMathRepaint(hwnd);
                }
                return 0;
            }

            if (state.active)
            {
                if (ch == L'\t') return 0; // Tab is handled in WM_KEYDOWN; block the WM_CHAR
                if (ch == L'=' && state.objectIndex < objects.size()) { 
                    // For system of equations, we allow typing the equals sign in equations (e.g., x+y=5)
                    // Only trigger calculation for non-system types
                    if (objects[state.objectIndex].type != MathType::SystemOfEquations) {
                        if (!mgr.CanCalculateResult(objects[state.objectIndex])) {
                            return 0;
                        }
                        // For other types (fractions, etc.), trigger calculation normally
                        ShowCaret(hwnd); 
                        state.active = false; 
                        TriggerCalculation(hwnd, state.objectIndex); 
                        return 0; 
                    }
                    // For system of equations, allow the equals sign to be typed normally
                    if (state.objectIndex < objects.size()) {
                        auto& obj = objects[state.objectIndex];
                        std::wstring& target = GetActiveEditText(obj, state.activePart, state.activeNodePath);
                        target.push_back(ch);
                        RefreshActiveSlotText(obj, state.activePart);
                        SyncFractionAnchorLength(hwnd, obj);
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                        UpdateResultIfPresent(hwnd, state.objectIndex);
                        RefreshUnitSuggestionPopup(hwnd);
                        RequestMathRepaint(hwnd);
                    }
                    return 0;
                }
                if (ch == L' ' && state.objectIndex < objects.size()) {
                    auto& obj = objects[state.objectIndex];
                    if (TryInsertNestedMathCommand(obj, state)) {
                        HideUnitSuggestionPopup(hwnd);
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                        UpdateResultIfPresent(hwnd, state.objectIndex);
                        RequestMathRepaint(hwnd);
                        return 0;
                    }
                }
                if (iswprint(ch) && ch != L'^' && ch != L'_') {
                    if (state.objectIndex < objects.size()) {
                        auto& obj = objects[state.objectIndex];
                        std::wstring& target = GetActiveEditText(obj, state.activePart, state.activeNodePath);
                        if (state.activeNodePath.empty() && state.activePart == 3 && target == L"{}") target = std::wstring(1, L'{') + ch + L'}';
                        else if (state.activeNodePath.empty() && state.activePart == 3 && target.size() >= 2 && target.front() == L'{' && target.back() == L'}') target.insert(target.size() - 1, 1, ch);
                        else target.push_back(ch);
                        RefreshActiveSlotText(obj, state.activePart);
                        SyncFractionAnchorLength(hwnd, obj);
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                        UpdateResultIfPresent(hwnd, state.objectIndex);
                        RefreshUnitSuggestionPopup(hwnd, ch == L'*' || ch == L'/');
                        RequestMathRepaint(hwnd);
                    }
                    return 0;
                }
                if (ch == L'^') { HideUnitSuggestionPopup(hwnd); state.activePart = 1; state.activeNodePath.clear(); RequestMathRepaint(hwnd); return 0; }
                if (ch == L'_') { HideUnitSuggestionPopup(hwnd); state.activePart = 2; state.activeNodePath.clear(); RequestMathRepaint(hwnd); return 0; }
                if (ch == 0x1B) {
                    if (!state.activeNodePath.empty()) {
                        state.activeNodePath.clear();
                        HideUnitSuggestionPopup(hwnd);
                        RequestMathRepaint(hwnd);
                        return 0;
                    }
                    HideUnitSuggestionPopup(hwnd); ShowCaret(hwnd); state.active = false; RequestMathRepaint(hwnd); return 0;
                }
                HideUnitSuggestionPopup(hwnd); ShowCaret(hwnd); state.active = false; state.activeNodePath.clear();
            }

            if ((ch == L' ' || ch == L'_' || ch == L'^') && !g_currentCommand.empty())
            {
                if (ch == L' ' && TryInsertFunctionTemplate(hwnd, (LONG)selEndBefore, g_currentCommand)) {
                    g_currentCommand.clear();
                    g_currentNumber.clear();
                    return 0;
                }

                MathObject obj; bool found = false;
                wchar_t anchorStr[10] = { 0 };
                LONG anchorLen = 5;
                if (g_currentCommand == L"\\sum") {
                    obj.type = MathType::Summation; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(L"N", L"i=0", L"{}"); found = true;
                } else if (g_currentCommand == L"\\prod") {
                    obj.type = MathType::Product; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(L"N", L"i=1", L"{}"); found = true;
                } else if (g_currentCommand == L"\\expr") {
                    obj.type = MathType::Sum; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(); found = true;
                } else if (g_currentCommand == L"\\expr") {
                    obj.type = MathType::Sum; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(); found = true;
                } else if (g_currentCommand == L"\\int") {
                    obj.type = MathType::Integral; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(L"b", L"a", L"{}"); found = true;
                } else if (g_currentCommand == L"\\sys") {
                    obj.type = MathType::SystemOfEquations; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(); found = true;
                } else if (g_currentCommand == L"\\frac") {
                    obj.type = MathType::Fraction; wcscpy_s(anchorStr, L"\u2500\u2500\u2500");
                    obj.SetParts(); obj.EnsureStructuredEditLeaf(1); obj.EnsureStructuredEditLeaf(2); anchorLen = 3; found = true;
                } else if (g_currentCommand == L"\\sqrt") {
                    obj.type = MathType::SquareRoot; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(); obj.EnsureStructuredEditLeaf(1); found = true;
                } else if (g_currentCommand == L"\\abs") {
                    obj.type = MathType::AbsoluteValue; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(); found = true;
                } else if (g_currentCommand == L"\\pow") {
                    obj.type = MathType::Power; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(); found = true;
                } else if (g_currentCommand == L"\\log") {
                    obj.type = MathType::Logarithm; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetParts(L"10"); found = true;
                } else if (g_currentCommand == L"\\mat") {
                    obj.type = MathType::Matrix; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetMatrix2x2(L"a", L"b", L"c", L"d");
                    obj.EnsureStructuredEditLeaf(1); obj.EnsureStructuredEditLeaf(2); obj.EnsureStructuredEditLeaf(3); obj.EnsureStructuredEditLeaf(4); found = true;
                } else if (g_currentCommand == L"\\det") {
                    obj.type = MathType::Determinant; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.SetMatrix2x2(L"a", L"b", L"c", L"d");
                    obj.EnsureStructuredEditLeaf(1); obj.EnsureStructuredEditLeaf(2); obj.EnsureStructuredEditLeaf(3); obj.EnsureStructuredEditLeaf(4); found = true;
                }

                if (found) {
                    const LONG cmdLen = (LONG)g_currentCommand.size();
                    const LONG cmdStart = (LONG)selEndBefore - cmdLen;
                    if (cmdStart >= 0) {
                        wchar_t v[64] = {0}; TEXTRANGEW tr = { {cmdStart, (LONG)selEndBefore}, v };
                        SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                        if (g_currentCommand == v) {
                            {
                                ScopedNoRedraw noRedraw(hwnd);
                                SendMessage(hwnd, EM_SETSEL, (WPARAM)cmdStart, (LPARAM)selEndBefore);
                                CHARFORMAT2W cf = {};
                                cf.cbSize = sizeof(cf);
                                SendMessage(hwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
                                LONG originalFontHeight = cf.yHeight;
                                // SquareRoot and AbsoluteValue: normal-height anchors (drawn with GDI lines)
                                // Others: 2x-height anchors to reserve vertical space for limits
                                if (obj.type != MathType::SquareRoot && obj.type != MathType::AbsoluteValue) {
                                    cf.dwMask |= CFM_SIZE; cf.yHeight *= 2;
                                    SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
                                }
                                SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)anchorStr);
                                // Push adjacent lines away from the math notation
                                {
                                    PARAFORMAT2 pf2 = {};
                                    pf2.cbSize = sizeof(pf2);
                                    pf2.dwMask = PFM_SPACEBEFORE | PFM_SPACEAFTER;
                                    if (obj.type == MathType::SquareRoot || obj.type == MathType::AbsoluteValue) {
                                        pf2.dySpaceBefore = (LONG)(originalFontHeight * 0.5);
                                        pf2.dySpaceAfter  = (LONG)(originalFontHeight * 0.3);
                                    } else {
                                        pf2.dySpaceBefore = (LONG)(originalFontHeight * 1.5);
                                        pf2.dySpaceAfter  = (LONG)(originalFontHeight * 1.5);
                                    }
                                    SendMessage(hwnd, EM_SETPARAFORMAT, 0, (LPARAM)&pf2);
                                }
                                mgr.ShiftObjectsAfter(selEndBefore, anchorLen - cmdLen);
                                obj.barStart = cmdStart; obj.barLen = anchorLen;
                                objects.push_back(std::move(obj));
                                MathObject& insertedObj = objects.back();
                                state.objectIndex = objects.size() - 1; if (!state.active) HideCaret(hwnd);
                                state.active = true;
                                state.activeNodePath.clear();
                                if (insertedObj.type == MathType::SystemOfEquations) state.activePart = 1;
                                else if (insertedObj.type == MathType::Fraction) state.activePart = (ch == L'_') ? 2 : 1;
                                else if (insertedObj.type == MathType::SquareRoot) state.activePart = (ch == L'_') ? 2 : 1;
                                else if (insertedObj.type == MathType::AbsoluteValue) state.activePart = 1;
                                else if (insertedObj.type == MathType::Power) state.activePart = 1;
                                else if (insertedObj.type == MathType::Logarithm) state.activePart = 2;  // Start with argument
                                else if (insertedObj.type == MathType::Sum) state.activePart = 1;  // Start with numbers to sum
                                else if (insertedObj.type == MathType::Matrix || insertedObj.type == MathType::Determinant) state.activePart = 1;
                                else if (ch == L'^') state.activePart = 1; else if (ch == L'_') state.activePart = 2; else state.activePart = 3;
                                if ((insertedObj.type == MathType::SquareRoot || insertedObj.type == MathType::Fraction) && insertedObj.EnterFirstStructuredLeaf(state.activePart, state.activeNodePath)) {
                                    SendMessage(hwnd, EM_SETSEL, (WPARAM)(cmdStart + anchorLen), (LPARAM)(cmdStart + anchorLen));
                                }
                                SendMessage(hwnd, EM_SETSEL, (WPARAM)(cmdStart + anchorLen), (LPARAM)(cmdStart + anchorLen));
                            }
                            HideUnitSuggestionPopup(hwnd);
                            RequestMathRepaint(hwnd); g_currentCommand.clear(); return 0;
                        }
                    }
                }
                g_currentCommand.clear();
            }

            if (ch == L'/' && !g_currentNumber.empty()) {
                const LONG nLen = (LONG)g_currentNumber.size();
                const LONG nS = (LONG)selEndBefore - nLen;
                const LONG bL = std::max<LONG>(3, nLen);
                {
                    ScopedNoRedraw noRedraw(hwnd);
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)nS, (LPARAM)selEndBefore);
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)bL, L'\u2500').c_str());
                    HideAnchorChars(hwnd, nS, bL);
                    mgr.ShiftObjectsAfter(selEndBefore, bL - nLen);
                    MathObject obj; obj.type = MathType::Fraction; obj.barStart = nS; obj.barLen = bL; obj.SetParts(g_currentNumber);
                    objects.push_back(std::move(obj));
                    state.objectIndex = objects.size() - 1; if (!state.active) HideCaret(hwnd);
                    state.active = true; state.activePart = 2;
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)(nS + bL), (LPARAM)(nS + bL));
                }
                HideUnitSuggestionPopup(hwnd);
                RequestMathRepaint(hwnd); g_currentNumber.clear(); return 0;
            }

            if (ch >= L'0' && ch <= L'9') { g_currentNumber += ch; g_currentCommand.clear(); }
            else if (ch == L'\\' || !g_currentCommand.empty()) { g_currentCommand += ch; g_currentNumber.clear(); }
            else { g_currentNumber.clear(); g_currentCommand.clear(); }

            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            const LONG delta = (LONG)GetWindowTextLengthW(hwnd) - lenBefore;
            const bool affectsObjects = (delta != 0) && HasObjectAtOrAfter(objects, (LONG)selEndBefore);
            mgr.ShiftObjectsAfter((LONG)selEndBefore, delta);
            if (affectsObjects) RequestMathRepaint(hwnd);
            return res;
        }
        }
        return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
    }
}

void ApplyMathEditInsets(HWND hRichEdit)
{
    if (!hRichEdit)
        return;

    RECT client = {};
    GetClientRect(hRichEdit, &client);

    RECT formatRect = client;
    formatRect.left += kMathEditInsetLeft;
    formatRect.top += kMathEditInsetTop;
    formatRect.right -= kMathEditInsetRight;
    formatRect.bottom -= kMathEditInsetBottom;

    if (formatRect.right < formatRect.left)
        formatRect.right = formatRect.left;
    if (formatRect.bottom < formatRect.top)
        formatRect.bottom = formatRect.top;

    SendMessage(hRichEdit, EM_SETRECTNP, 0, (LPARAM)&formatRect);
    RedrawWindow(hRichEdit, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
}

bool InstallMathSupport(HWND hRichEdit)
{
    if (!hRichEdit) return false;
    g_hEdit = hRichEdit; g_currentNumber.clear(); g_currentCommand.clear();
    ClearUnitSuggestionPopup();
    MathManager::Get().Clear();
    if (!g_originalProc) g_originalProc = (WNDPROC)SetWindowLongPtr(g_hEdit, GWLP_WNDPROC, (LONG_PTR)MathRichEditProc);
    ApplyMathEditInsets(g_hEdit);
    return !!g_originalProc;
}

void ResetMathSupport() { g_currentNumber.clear(); g_currentCommand.clear(); ClearUnitSuggestionPopup(); MathManager::Get().Clear(); if (g_hEdit) RedrawWindow(g_hEdit, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE); }

bool DebugGetUnitSuggestionState(std::vector<std::wstring>& outSuggestions, size_t& outSelectedIndex, std::wstring* outPrefix, RECT* outPopupRect)
{
    outSuggestions = g_unitSuggestionPopup.items;
    outSelectedIndex = g_unitSuggestionPopup.selectedIndex;
    if (outPrefix)
        *outPrefix = g_unitSuggestionPopup.prefix;
    if (outPopupRect)
        *outPopupRect = g_unitSuggestionPopup.popupRect;
    return g_unitSuggestionPopup.visible;
}

void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator)
{
    if (!hEdit) return;
    DWORD s, e; SendMessage(hEdit, EM_GETSEL, (WPARAM)&s, (LPARAM)&e);
    const LONG bL = (LONG)std::max<size_t>(3, std::max(numerator.size(), denominator.size()));
    SendMessage(hEdit, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)bL, L'\u2500').c_str());
    HideAnchorChars(hEdit, (LONG)s, bL);
    MathObject obj; obj.type = MathType::Fraction; obj.barStart = (LONG)s; obj.barLen = bL; obj.SetParts(numerator, denominator);
    MathManager::Get().GetObjects().push_back(std::move(obj));
    MathManager::Get().ShiftObjectsAfter((LONG)s + 1, bL - (LONG)(e - s));
    SendMessage(hEdit, EM_SETSEL, (WPARAM)(s + bL), (LPARAM)(s + bL));
    RedrawWindow(hEdit, nullptr, nullptr, RDW_INVALIDATE | RDW_NOERASE);
}

bool SerializeMathDocument(HWND hEdit, std::wstring& outPayload)
{
    outPayload.clear();
    if (!hEdit)
        return false;

    MathDocumentSnapshot snapshot;
    if (!BuildMathDocumentSnapshot(hEdit, snapshot))
        return false;

    outPayload = SerializeMathDocumentSnapshot(snapshot);
    return true;
}

bool TryDeserializeMathDocument(HWND hEdit, const std::wstring& payload)
{
    if (!hEdit)
        return false;

    MathDocumentSnapshot snapshot;
    if (!TryDeserializeMathDocumentSnapshot(payload, snapshot))
        return false;

    return RestoreMathDocumentSnapshot(hEdit, snapshot);
}

bool DebugRunStructuredRoundTripSelfTest(HWND hEdit, std::wstring& outDetails)
{
    outDetails.clear();
    if (!hEdit)
    {
        outDetails = L"No RichEdit handle";
        return false;
    }

    const int originalLen = GetWindowTextLengthW(hEdit);
    std::wstring originalText;
    if (originalLen > 0)
    {
        std::vector<wchar_t> buffer((size_t)originalLen + 1, L'\0');
        GetWindowTextW(hEdit, buffer.data(), originalLen + 1);
        originalText.assign(buffer.data());
    }

    SetWindowTextW(hEdit, L"");
    ResetMathSupport();
    SendMessage(hEdit, EM_SETSEL, 0, 0);

    SendChars(hEdit, L"\\sqrt");
    SendMessage(hEdit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(hEdit, L"9+\\frac");
    SendMessage(hEdit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(hEdit, L"16");
    SendMessage(hEdit, WM_KEYDOWN, VK_TAB, 0);
    SendChars(hEdit, L"4");

    auto& objects = MathManager::Get().GetObjects();
    if (objects.size() != 1)
    {
        outDetails = L"Expected one structured object before round-trip.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    const std::wstring originalPayload = objects[0].SerializeTransferPayload();
    MathObject duplicate;
    if (!MathObject::TryDeserializeTransferPayload(originalPayload, duplicate))
    {
        outDetails = L"Failed to deserialize the structured transfer payload during self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    const LONG insertPos = objects[0].barStart + objects[0].barLen;
    SendMessage(hEdit, EM_SETSEL, (WPARAM)insertPos, (LPARAM)insertPos);
    if (!InsertStructuredObjectAtSelection(hEdit, std::move(duplicate)))
    {
        outDetails = L"Failed to reinsert structured object from transfer payload.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    if (objects.size() != 2)
    {
        outDetails = L"Expected two structured objects after round-trip reinsertion.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    if (objects[1].SerializeTransferPayload() != originalPayload)
    {
        outDetails = L"Round-tripped object payload did not match the original.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    SetWindowTextW(hEdit, originalText.c_str());
    ResetMathSupport();
    return true;
}

bool DebugRunStructuredFragmentRoundTripSelfTest(HWND hEdit, std::wstring& outDetails)
{
    outDetails.clear();
    if (!hEdit)
    {
        outDetails = L"No RichEdit handle";
        return false;
    }

    const int originalLen = GetWindowTextLengthW(hEdit);
    std::wstring originalText;
    if (originalLen > 0)
    {
        std::vector<wchar_t> buffer((size_t)originalLen + 1, L'\0');
        GetWindowTextW(hEdit, buffer.data(), originalLen + 1);
        originalText.assign(buffer.data());
    }

    SetWindowTextW(hEdit, L"");
    ResetMathSupport();
    SendMessage(hEdit, EM_SETSEL, 0, 0);

    SendChars(hEdit, L"A ");
    SendChars(hEdit, L"3/4");
    SendMessage(hEdit, WM_KEYDOWN, VK_RETURN, 0);
    SendChars(hEdit, L" B ");
    SendChars(hEdit, L"\\sqrt");
    SendMessage(hEdit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(hEdit, L"16");
    SendMessage(hEdit, WM_KEYDOWN, VK_RETURN, 0);
    SendChars(hEdit, L" C");

    auto& objects = MathManager::Get().GetObjects();
    if (objects.size() != 2)
    {
        outDetails = L"Expected two structured objects before fragment round-trip.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    const LONG firstObjectStart = std::min(objects[0].barStart, objects[1].barStart);
    const LONG secondObjectEnd = std::max(objects[0].barStart + objects[0].barLen, objects[1].barStart + objects[1].barLen);
    SendMessage(hEdit, EM_SETSEL, (WPARAM)firstObjectStart, (LPARAM)secondObjectEnd);
    LONG selStart = 0;
    LONG selEnd = 0;
    std::vector<size_t> objectIndices;
    if (!TryCollectSelectedMathObjects(hEdit, selStart, selEnd, objectIndices) || objectIndices.size() != 2)
    {
        outDetails = L"Expected a selection containing two fully enclosed math objects for fragment self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    MathClipboardFragment fragment;
    std::wstring plainText;
    if (!BuildClipboardFragmentFromSelection(hEdit, selStart, selEnd, objectIndices, fragment, plainText))
    {
        outDetails = L"Failed to build structured fragment payload during self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    const std::wstring originalFragmentPayload = SerializeClipboardFragment(fragment);
    const LONG fullLen = GetWindowTextLengthW(hEdit);
    SendMessage(hEdit, EM_SETSEL, (WPARAM)fullLen, (LPARAM)fullLen);
    if (!OpenClipboard(hEdit))
    {
        outDetails = L"Failed to open clipboard during fragment self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }
    EmptyClipboard();
    const bool clipboardOk = SetClipboardPayload(GetMathFragmentClipboardFormat(), originalFragmentPayload) && SetClipboardUnicodeText(plainText);
    CloseClipboard();
    if (!clipboardOk || !TryPasteMathFragmentFromClipboard(hEdit))
    {
        outDetails = L"Failed to paste structured fragment payload during self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    const LONG newLen = GetWindowTextLengthW(hEdit);
    const LONG pastedSelEnd = fullLen + (LONG)fragment.rawText.size();
    SendMessage(hEdit, EM_SETSEL, (WPARAM)fullLen, (LPARAM)std::min(newLen, pastedSelEnd));
    std::vector<size_t> duplicatedIndices;
    if (!TryCollectSelectedMathObjects(hEdit, selStart, selEnd, duplicatedIndices) || duplicatedIndices.size() != 2)
    {
        outDetails = L"Expected duplicated structured fragment to contain two math objects after paste.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    MathClipboardFragment duplicatedFragment;
    std::wstring duplicatedPlainText;
    if (!BuildClipboardFragmentFromSelection(hEdit, selStart, selEnd, duplicatedIndices, duplicatedFragment, duplicatedPlainText))
    {
        outDetails = L"Failed to rebuild duplicated fragment payload during self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    if (SerializeClipboardFragment(duplicatedFragment) != originalFragmentPayload)
    {
        outDetails = L"Round-tripped fragment payload did not match the original.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    SetWindowTextW(hEdit, originalText.c_str());
    ResetMathSupport();
    return true;
}

bool DebugRunStructuredDocumentRoundTripSelfTest(HWND hEdit, std::wstring& outDetails)
{
    outDetails.clear();
    if (!hEdit)
    {
        outDetails = L"No RichEdit handle";
        return false;
    }

    const int originalLen = GetWindowTextLengthW(hEdit);
    std::wstring originalText;
    if (originalLen > 0)
    {
        std::vector<wchar_t> buffer((size_t)originalLen + 1, L'\0');
        GetWindowTextW(hEdit, buffer.data(), originalLen + 1);
        originalText.assign(buffer.data());
    }

    SetWindowTextW(hEdit, L"");
    ResetMathSupport();
    SendMessage(hEdit, EM_SETSEL, 0, 0);

    SendChars(hEdit, L"Header ");
    SendChars(hEdit, L"3/4");
    SendMessage(hEdit, WM_KEYDOWN, VK_RETURN, 0);
    SendChars(hEdit, L"Middle ");
    SendChars(hEdit, L"\\sqrt");
    SendMessage(hEdit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(hEdit, L"9+\\frac");
    SendMessage(hEdit, WM_CHAR, (WPARAM)L' ', 0);
    SendChars(hEdit, L"16");
    SendMessage(hEdit, WM_KEYDOWN, VK_TAB, 0);
    SendChars(hEdit, L"4");
    SendMessage(hEdit, WM_KEYDOWN, VK_RETURN, 0);
    SendChars(hEdit, L" Footer");

    auto& objects = MathManager::Get().GetObjects();
    if (objects.size() != 2)
    {
        outDetails = L"Expected two structured objects before document round-trip.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    std::wstring originalPayload;
    if (!SerializeMathDocument(hEdit, originalPayload))
    {
        outDetails = L"Failed to serialize structured document snapshot during self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    if (!TryDeserializeMathDocument(hEdit, originalPayload))
    {
        outDetails = L"Failed to restore structured document snapshot during self-test.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    if (MathManager::Get().GetObjects().size() != 2)
    {
        outDetails = L"Expected two structured objects after document snapshot restore.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    std::wstring roundTrippedPayload;
    if (!SerializeMathDocument(hEdit, roundTrippedPayload))
    {
        outDetails = L"Failed to reserialize structured document snapshot after restore.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    if (roundTrippedPayload != originalPayload)
    {
        outDetails = L"Round-tripped document snapshot did not match the original.";
        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return false;
    }

    SetWindowTextW(hEdit, originalText.c_str());
    ResetMathSupport();
    return true;
}
