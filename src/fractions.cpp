#ifndef NOMINMAX
#define NOMINMAX
#endif

#include "fractions.h"

#include <algorithm>
#include <vector>

namespace
{
    static const bool kDebugOverlay = false;

    struct FractionTypingState
    {
        bool active = false;      // typing denominator or numerator
        bool isNumerator = false; // true if editing numerator, false for denominator
        size_t fractionIndex = 0; // index into fractions
    };

    struct FractionSpan
    {
        LONG barStart = 0;
        LONG barLen = 0;
        std::wstring numerator;
        std::wstring denominator;
    };

    HWND g_hEdit = nullptr;
    WNDPROC g_originalProc = nullptr;

    std::wstring g_currentNumber;
    FractionTypingState g_state;
    std::vector<FractionSpan> g_fractions;

    static void ShiftFractionsAfter(LONG atPosInclusive, LONG delta)
    {
        if (delta == 0)
            return;
        for (auto& f : g_fractions)
        {
            if (f.barStart >= atPosInclusive)
                f.barStart += delta;
        }
    }

    static bool IsPosInsideAnyFractionBar(LONG pos, size_t* outIndex = nullptr)
    {
        for (size_t i = 0; i < g_fractions.size(); ++i)
        {
            const auto& f = g_fractions[i];
            // Bar covers character indices [barStart, barStart + barLen).
            if (pos >= f.barStart && pos < (f.barStart + f.barLen))
            {
                if (outIndex)
                    *outIndex = i;
                return true;
            }
        }
        return false;
    }

    static void UpdateResultIfPresent(HWND hwnd, size_t fracIdx)
    {
        if (fracIdx >= g_fractions.size()) return;
        auto& f = g_fractions[fracIdx];
        
        // Search if " = " or " \uFF1D " follows the fraction
        const LONG startPos = f.barStart + f.barLen + (LONG)f.denominator.size();
        
        const LONG totalLen = (LONG)GetWindowTextLengthW(hwnd);
        bool foundEquals = false;
        LONG equalsPos = -1;

        // Search for the equals sign in the next few characters
        for (int i = 0; i < 15 && (startPos + i) < totalLen; ++i)
        {
            wchar_t buf[2] = {0, 0};
            TEXTRANGEW tr = {};
            tr.chrg.cpMin = startPos + i;
            tr.chrg.cpMax = startPos + i + 1;
            tr.lpstrText = buf;
            SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr);

            if (buf[0] == L'\uFF1D' || buf[0] == L'=')
            {
                foundEquals = true;
                equalsPos = startPos + i;
                break;
            }
        }

        if (foundEquals)
        {
            // Calculate new result
            if (f.numerator.empty() || f.denominator.empty()) return;
            double numVal = _wtof(f.numerator.c_str());
            double denVal = _wtof(f.denominator.c_str());
            if (denVal == 0.0) return;
            double result = numVal / denVal;
            
            wchar_t resultBuf[128];
            if (result == (long long)result)
                swprintf_s(resultBuf, L" %lld", (long long)result);
            else
                swprintf_s(resultBuf, L" %.3g", result);

            // Find current number until space or end of line
            LONG searchStart = equalsPos + 1;
            // Skip space after = if present
            {
                wchar_t buf[2] = {0, 0};
                TEXTRANGEW tr = {{searchStart, searchStart + 1}, buf};
                if (SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr) > 0 && buf[0] == L' ')
                    searchStart++;
            }

            LONG searchEnd = searchStart;
            while (searchEnd < totalLen)
            {
                wchar_t buf[2] = {0, 0};
                TEXTRANGEW tr2 = { {searchEnd, searchEnd + 1}, buf };
                if (SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr2) == 0) break;
                if (!iswdigit(buf[0]) && buf[0] != L'.' && buf[0] != L'-') break;
                searchEnd++;
            }

            // Save selection to restore it
            DWORD selStart = 0, selEnd = 0;
            SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

            // Replace current result
            SendMessage(hwnd, EM_SETSEL, (WPARAM)searchStart, (LPARAM)searchEnd);
            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)resultBuf);
            
            LONG delta = (LONG)wcslen(resultBuf) - (searchEnd - searchStart);
            ShiftFractionsAfter(searchStart, delta);
            
            // Restore selection
            SendMessage(hwnd, EM_SETSEL, (WPARAM)selStart, (LPARAM)selEnd);
            InvalidateRect(hwnd, nullptr, FALSE);
        }
    }

    static void TriggerCalculation(HWND hwnd, size_t fracIdx)
    {
        if (fracIdx >= g_fractions.size()) return;
        auto& f = g_fractions[fracIdx];
        
        if (f.numerator.empty() || f.denominator.empty()) return;

        double numVal = _wtof(f.numerator.c_str());
        double denVal = _wtof(f.denominator.c_str());
        
        if (denVal == 0.0) return;

        double result = numVal / denVal;
        
        wchar_t resultBuf[128];
        if (result == (long long)result)
            swprintf_s(resultBuf, L" %lld", (long long)result);
        else
            swprintf_s(resultBuf, L" %.3g", result);

        const LONG insertPos = f.barStart + f.barLen + (LONG)f.denominator.size();
        SendMessage(hwnd, EM_SETSEL, (WPARAM)insertPos, (LPARAM)insertPos);

        // Save current format
        CHARFORMAT2W baseCf = {};
        baseCf.cbSize = sizeof(baseCf);
        SendMessage(hwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&baseCf);

        // 1. Insert " = "
        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L" \uFF1D ");

        // 2. Format the " = " part
        SendMessage(hwnd, EM_SETSEL, (WPARAM)insertPos, (LPARAM)(insertPos + 3));
        
        CHARFORMAT2W cf = baseCf;
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_OFFSET | CFM_COLOR | CFM_BOLD | CFM_SIZE | CFM_FACE;
        cf.yOffset = -40; // Approx 2 pixels
        cf.yHeight = (baseCf.yHeight * 125) / 100;
        cf.crTextColor = RGB(0, 102, 204);
        cf.dwEffects = CFE_BOLD;
        wcscpy_s(cf.szFaceName, L"Cambria Math");
        SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);

        // 3. Insert the result number
        const LONG resultPos = insertPos + 3;
        SendMessage(hwnd, EM_SETSEL, (WPARAM)resultPos, (LPARAM)resultPos);
        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)resultBuf);

        // 4. Reset format
        const LONG resultEnd = resultPos + (LONG)wcslen(resultBuf);
        SendMessage(hwnd, EM_SETSEL, (WPARAM)resultPos, (LPARAM)resultEnd);
        cf = baseCf;
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_OFFSET | CFM_COLOR | CFM_BOLD | CFM_SIZE | CFM_FACE;
        SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);

        // 5. Place caret at end
        SendMessage(hwnd, EM_SETSEL, (WPARAM)resultEnd, (LPARAM)resultEnd);

        ShiftFractionsAfter(insertPos, 3 + (LONG)wcslen(resultBuf));
        InvalidateRect(hwnd, nullptr, FALSE);
    }

    static bool TryGetCharPos(HWND hEdit, LONG charIndex, POINT& outPt)
    {
        // First try the standard EM_POSFROMCHAR with POINTL
        POINTL ptL = {};
        LRESULT r = SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)&ptL, (LPARAM)charIndex);
        if (r != -1)
        {
            outPt.x = (int)ptL.x;
            outPt.y = (int)ptL.y;
            return true;
        }

        // Fallback: legacy edit control packed x/y form
        LRESULT xy = SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)charIndex, 0);
        if (xy != -1)
        {
            outPt.x = (short)LOWORD(xy);
            outPt.y = (short)HIWORD(xy);
            return true;
        }

        return false;
    }

    static double ComputeRenderScaleFromBarSpan(HWND hEdit, HDC hdc, const FractionSpan& f, HFONT unzoomedBaseFont)
    {
        if (f.barLen <= 0 || !unzoomedBaseFont)
            return 1.0;

        LONG span = std::min<LONG>(f.barLen, 8);
        while (span >= 1)
        {
            POINT p0 = {}, pN = {};
            if (TryGetCharPos(hEdit, f.barStart, p0) && TryGetCharPos(hEdit, f.barStart + span, pN))
            {
                if (p0.y != pN.y)
                {
                    --span;
                    continue;
                }

                const int actualWidthPx = pN.x - p0.x;
                if (actualWidthPx > 0)
                {
                    HFONT old = (HFONT)SelectObject(hdc, unzoomedBaseFont);
                    const wchar_t barChar = L'\u2500';
                    SIZE one = {};
                    GetTextExtentPoint32W(hdc, &barChar, 1, &one);
                    SelectObject(hdc, old);

                    if (one.cx > 0)
                    {
                        const int expectedWidthPx = one.cx * (int)span;
                        if (expectedWidthPx > 0)
                            return (double)actualWidthPx / (double)expectedWidthPx;
                    }
                }
            }
            --span;
        }

        return 1.0;
    }

    static HFONT CreateScaledFontWithScale(HFONT baseFont, double scale, int percent)
    {
        LOGFONTW lf = {};
        if (!baseFont || GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
        {
            return CreateFontW(-11, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
        }

        if (!(scale > 0.0))
            scale = 1.0;

        const int sign = (lf.lfHeight < 0) ? -1 : 1;
        const int absH = (lf.lfHeight < 0) ? -lf.lfHeight : lf.lfHeight;
        const double scaled = (double)absH * scale * ((double)percent / 100.0);
        const int newAbsH = std::max(1, (int)(scaled + 0.5));
        lf.lfHeight = sign * newAbsH;
        return CreateFontIndirectW(&lf);
    }

    static COLORREF GetRichEditDefaultTextColor(HWND hEdit)
    {
        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        SendMessage(hEdit, EM_GETCHARFORMAT, (WPARAM)SCF_DEFAULT, (LPARAM)&cf);

        if ((cf.dwMask & CFM_COLOR) == 0 || (cf.dwEffects & CFE_AUTOCOLOR) != 0)
            return GetSysColor(COLOR_WINDOWTEXT);
        return cf.crTextColor;
    }


    static bool GetHitFractionPart(HWND hEdit, HDC hdc, POINT ptMouse, size_t* outIndex, bool* outIsNumerator)
    {
        for (size_t i = 0; i < g_fractions.size(); ++i)
        {
            const auto& f = g_fractions[i];
            
            POINT ptStart = {}, ptEnd = {};
            if (!TryGetCharPos(hEdit, f.barStart, ptStart)) continue;
            if (!TryGetCharPos(hEdit, f.barStart + std::max<LONG>(0, f.barLen - 1), ptEnd)) continue;

            HFONT baseFont = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
            if (!baseFont) baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

            const double renderScale = ComputeRenderScaleFromBarSpan(hEdit, hdc, f, baseFont);
            HFONT numFont = CreateScaledFontWithScale(baseFont, renderScale, 70);
            HFONT denFont = CreateScaledFontWithScale(baseFont, renderScale, 70);
            HFONT renderBaseFont = CreateScaledFontWithScale(baseFont, renderScale, 100);

            TEXTMETRICW tmBase = {};
            HFONT old = (HFONT)SelectObject(hdc, renderBaseFont);
            GetTextMetricsW(hdc, &tmBase);

            // We calculate xCenter by finding the start of the first char and the end 
            // of the last char in the bar.
            const int barWidth = (ptEnd.x - ptStart.x) + (int)(tmBase.tmAveCharWidth * renderScale);
            const int xCenter = ptStart.x + (barWidth / 2);

            // Anchor yMid to the font baseline. This is exactly where the 
            // Unicode horizontal bar character normally sits in the line.
            const int yMid = ptStart.y + tmBase.tmAscent - (int)(2 * renderScale);
            const int gap = std::max<int>(2, (int)(tmBase.tmHeight / 10));

            bool hit = false;
            // Numerator check
            {
                SelectObject(hdc, numFont);
                TEXTMETRICW tmNum = {};
                GetTextMetricsW(hdc, &tmNum);
                SIZE numSize = {};
                GetTextExtentPoint32W(hdc, f.numerator.c_str(), (int)f.numerator.size(), &numSize);

                int w = std::max<int>(30, numSize.cx);
                int h = std::max<int>(20, tmNum.tmHeight);

                RECT rcNum = { 
                    xCenter - (w / 2) - 15, 
                    yMid - gap - h - 10,
                    xCenter + (w / 2) + 15,
                    yMid - gap + 5
                };
                if (PtInRect(&rcNum, ptMouse)) {
                    *outIndex = i;
                    *outIsNumerator = true;
                    hit = true;
                }
            }

            if (!hit)
            {
                // Denominator check
                SelectObject(hdc, denFont);
                TEXTMETRICW tmDen = {};
                GetTextMetricsW(hdc, &tmDen);
                SIZE denSize = {};
                GetTextExtentPoint32W(hdc, f.denominator.c_str(), (int)f.denominator.size(), &denSize);

                int w = std::max<int>(30, denSize.cx);
                int h = std::max<int>(20, tmDen.tmHeight);

                RECT rcDen = {
                    xCenter - (w / 2) - 15,
                    yMid + gap - 5,
                    xCenter + (w / 2) + 15,
                    yMid + gap + h + 10
                };
                if (PtInRect(&rcDen, ptMouse)) {
                    *outIndex = i;
                    *outIsNumerator = false;
                    hit = true;
                }
            }

            SelectObject(hdc, old);
            DeleteObject(renderBaseFont);
            DeleteObject(numFont);
            DeleteObject(denFont);
            if (hit) return true;
        }
        return false;
    }

    static void DrawFractionOverBar(HWND hEdit, HDC hdc, const FractionSpan& f, size_t fIndex)
    {
        if (f.barLen <= 0)
            return;

        POINT ptStart = {}, ptEnd = {};
        if (!TryGetCharPos(hEdit, f.barStart, ptStart))
            return;
        if (!TryGetCharPos(hEdit, f.barStart + std::max<LONG>(0, f.barLen - 1), ptEnd))
            return;

        // Ensure we have valid coordinates
        if (ptStart.x == 0 && ptStart.y == 0 && ptEnd.x == 0 && ptEnd.y == 0)
        {
            // Fallback to center of window for debugging
            RECT rc;
            GetClientRect(hEdit, &rc);
            ptStart.x = rc.right / 2 - 50;
            ptStart.y = rc.bottom / 2 - 20;
            ptEnd.x = ptStart.x + 100;
            ptEnd.y = ptStart.y;
        }

        HFONT baseFont = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
        bool deleteBaseFont = false;
        if (!baseFont)
        {
            baseFont = CreateFontW(-16, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
            deleteBaseFont = true;
        }

        const double renderScale = ComputeRenderScaleFromBarSpan(hEdit, hdc, f, baseFont);

        HFONT renderBaseFont = CreateScaledFontWithScale(baseFont, renderScale, 100);
        HFONT numFont = CreateScaledFontWithScale(baseFont, renderScale, 70);
        HFONT denFont = CreateScaledFontWithScale(baseFont, renderScale, 70);

        const int saved = SaveDC(hdc);
        SelectClipRgn(hdc, nullptr);

        SetBkMode(hdc, TRANSPARENT);
        SetTextAlign(hdc, TA_BASELINE | TA_CENTER); 

        HFONT oldFont = (HFONT)SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmBase = {};
        GetTextMetricsW(hdc, &tmBase);

        const int barWidth = (ptEnd.x - ptStart.x) + (int)(tmBase.tmAveCharWidth);
        const int xCenter = ptStart.x + (barWidth / 2);

        // Anchor yMid to the baseline (tmAscent) of the RichEdit line.
        // This is exactly where the Unicode bar character sits.
        // Using a percentage of ascent (68%) ensures it aligns with math operators 
        // across different zoom levels and font sizes.
        const int yMid = ptStart.y + (int)(tmBase.tmAscent * 68 / 100);
        const int gap = std::max<int>(1, (int)(tmBase.tmHeight / 15));

        const COLORREF normalColor = RGB(0, 0, 0);
        const COLORREF activeColor = RGB(0, 102, 204);

        // Draw numerator
        {
            SelectObject(hdc, numFont);
            TEXTMETRICW tmNum = {};
            GetTextMetricsW(hdc, &tmNum);

            // Numerator baseline should be gap away from the bar center
            int yNumBaseline = yMid - gap;

            if (g_state.active && g_state.fractionIndex == fIndex && g_state.isNumerator)
                SetTextColor(hdc, activeColor);
            else
                SetTextColor(hdc, normalColor);

            // Draw placeholder if empty while editing
            if (f.numerator.empty() && g_state.active && g_state.fractionIndex == fIndex && g_state.isNumerator)
                TextOutW(hdc, xCenter, yNumBaseline, L"?", 1);
            else
                TextOutW(hdc, xCenter, yNumBaseline, f.numerator.c_str(), (int)f.numerator.size());
        }

        // Draw denominator
        {
            SelectObject(hdc, denFont);
            TEXTMETRICW tmDen = {};
            GetTextMetricsW(hdc, &tmDen);

            // Denominator top should be gap away from the bar center
            // Baseline is top + tmAscent.
            int yDenBaseline = yMid + gap + tmDen.tmAscent;

            if (g_state.active && g_state.fractionIndex == fIndex && !g_state.isNumerator)
                SetTextColor(hdc, activeColor);
            else
                SetTextColor(hdc, normalColor);

            if (f.denominator.empty() && g_state.active && g_state.fractionIndex == fIndex && !g_state.isNumerator)
                TextOutW(hdc, xCenter, yDenBaseline, L"?", 1);
            else
                TextOutW(hdc, xCenter, yDenBaseline, f.denominator.c_str(), (int)f.denominator.size());
        }

        SelectObject(hdc, oldFont);
        RestoreDC(hdc, saved);

        DeleteObject(renderBaseFont);
        DeleteObject(numFont);
        DeleteObject(denFont);
        if (deleteBaseFont) DeleteObject(baseFont);

        // Update title with debug info
        wchar_t buf[256];
        swprintf_s(buf, L"Fractions: %d | Editing: %s (%s)",
            (int)g_fractions.size(), 
            g_state.active ? L"YES" : L"NO", 
            g_state.active ? (g_state.isNumerator ? L"Num" : L"Den") : L"None");
        SetWindowTextW(GetParent(hEdit), buf);
    }

    static LRESULT CALLBACK FractionRichEditProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
    {
        switch (uMsg)
        {
        case WM_SETCURSOR:
        {
            if (LOWORD(lParam) == HTCLIENT)
            {
                POINT pt;
                GetCursorPos(&pt);
                ScreenToClient(hwnd, &pt);

                HDC hdc = GetDC(hwnd);
                size_t idx = 0;
                bool isNum = false;
                bool hit = GetHitFractionPart(hwnd, hdc, pt, &idx, &isNum);
                
                // Fallback: Check if over bar characters directly
                if (!hit)
                {
                    POINTL ptl = { pt.x, pt.y };
                    LRESULT charIdx = SendMessage(hwnd, EM_CHARFROMPOS, 0, (LPARAM)&ptl);
                    size_t dummy;
                    if (IsPosInsideAnyFractionBar((LONG)charIdx, &dummy))
                        hit = true;
                }
                
                ReleaseDC(hwnd, hdc);

                if (hit)
                {
                    SetCursor(LoadCursor(nullptr, (LPCWSTR)IDC_HAND));
                    return TRUE;
                }
            }
            break;
        }

        case WM_SETFOCUS:
        {
            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            if (g_state.active) HideCaret(hwnd);
            return res;
        }

        case WM_LBUTTONDOWN:
        {
            POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
            HDC hdc = GetDC(hwnd);
            size_t idx = 0;
            bool isNum = false;
            bool hit = GetHitFractionPart(hwnd, hdc, pt, &idx, &isNum);
            ReleaseDC(hwnd, hdc);

            if (hit)
            {
                SetFocus(hwnd); // Ensure we have focus to receive characters
                if (!g_state.active) HideCaret(hwnd);
                g_state.active = true;
                g_state.fractionIndex = idx;
                g_state.isNumerator = isNum;
                g_currentNumber.clear(); 
                
                // Move RichEdit caret to the start of the bar so typing 
                // feels "anchored" to the fraction location internally.
                const auto& f = g_fractions[idx];
                SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)f.barStart);
                
                InvalidateRect(hwnd, nullptr, FALSE);
                return 0; 
            }
            else
            {
                if (g_state.active)
                {
                    ShowCaret(hwnd);
                    g_state.active = false;
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
            }
            break; 
        }

        case WM_LBUTTONUP:
        {
            POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
            HDC hdc = GetDC(hwnd);
            size_t idx = 0;
            bool isNum = false;
            if (GetHitFractionPart(hwnd, hdc, pt, &idx, &isNum))
            {
                ReleaseDC(hwnd, hdc);
                return 0; // Swallow up message too
            }
            ReleaseDC(hwnd, hdc);
            break;
        }

        case WM_PRINTCLIENT:
        {
            const LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            HDC hdc = (HDC)wParam;
            if (hdc && (!g_fractions.empty() || kDebugOverlay))
            {
                const int saved = SaveDC(hdc);
                SelectClipRgn(hdc, nullptr);
                for (size_t i = 0; i < g_fractions.size(); ++i)
                    DrawFractionOverBar(hwnd, hdc, g_fractions[i], i);
                RestoreDC(hdc, saved);
            }
            return res;
        }

        case WM_MOUSEWHEEL:
        {
            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            if (!g_fractions.empty())
                InvalidateRect(hwnd, nullptr, FALSE);
            return res;
        }

        case WM_PAINT:
        {
            // Let RichEdit do its normal paint first.
            const LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);

            // Use GetDC instead of BeginPaint here because the original proc 
            // already validated the window's update region.
            HDC hdc = GetDC(hwnd);
            if (hdc)
            {
                if (!g_fractions.empty() || kDebugOverlay)
                {
                    const int saved = SaveDC(hdc);
                    SelectClipRgn(hdc, nullptr);
                    for (size_t i = 0; i < g_fractions.size(); ++i)
                        DrawFractionOverBar(hwnd, hdc, g_fractions[i], i);
                    RestoreDC(hdc, saved);
                }
                ReleaseDC(hwnd, hdc);
            }

            return res;
        }

        case WM_MOUSEMOVE:
        {
            // Set the cursor here too to ensure it feels responsive
            POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
            HDC hdc = GetDC(hwnd);
            size_t idx = 0;
            bool isNum = false;
            bool hit = GetHitFractionPart(hwnd, hdc, pt, &idx, &isNum);
            
            if (!hit)
            {
                POINTL ptl = { pt.x, pt.y };
                LRESULT charIdx = SendMessage(hwnd, EM_CHARFROMPOS, 0, (LPARAM)&ptl);
                size_t dummy;
                if (IsPosInsideAnyFractionBar((LONG)charIdx, &dummy))
                    hit = true;
            }
            
            ReleaseDC(hwnd, hdc);

            if (hit)
            {
                SetCursor(LoadCursor(nullptr, (LPCWSTR)IDC_HAND));
            }

            return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
        }

        case WM_KEYDOWN:
        {
            const LONG lenBefore = (LONG)GetWindowTextLengthW(hwnd);
            DWORD selStart = 0, selEnd = 0;
            SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

            if (wParam == VK_RETURN)
            {
                if (g_state.active)
                {
                    g_state.active = false;
                    UpdateResultIfPresent(hwnd, g_state.fractionIndex);
                    ShowCaret(hwnd);
                    InvalidateRect(hwnd, nullptr, FALSE);
                    return 0; 
                }
                
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                const LONG lenAfter = (LONG)GetWindowTextLengthW(hwnd);
                ShiftFractionsAfter((LONG)selEnd, lenAfter - lenBefore);
                return res;
            }

            if (wParam == VK_BACK)
            {
                if (g_state.active)
                {
                    // For backspace in active mode, we want WM_CHAR to handle 
                    // character-by-character numerator/denominator deletion.
                    return 0; 
                }

                size_t fracIndex = 0;
                if (IsPosInsideAnyFractionBar((LONG)selEnd - 1, &fracIndex))
                {
                    auto f = g_fractions[fracIndex];
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)(f.barStart + f.barLen));
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                    g_fractions.erase(g_fractions.begin() + (ptrdiff_t)fracIndex);
                    ShiftFractionsAfter(f.barStart + 1, -f.barLen);
                    InvalidateRect(hwnd, nullptr, FALSE);
                    return 0;
                }

                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                const LONG lenAfter = (LONG)GetWindowTextLengthW(hwnd);
                ShiftFractionsAfter((LONG)selEnd, lenAfter - lenBefore);
                return res;
            }

            if (wParam == VK_DELETE)
            {
                const LONG lenBeforeDel = (LONG)GetWindowTextLengthW(hwnd);
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                const LONG lenAfterDel = (LONG)GetWindowTextLengthW(hwnd);
                if (lenAfterDel != lenBeforeDel)
                {
                    ShiftFractionsAfter((LONG)selEnd, lenAfterDel - lenBeforeDel);
                    InvalidateRect(hwnd, nullptr, FALSE);
                }

                if (g_state.active)
                {
                    ShowCaret(hwnd);
                    g_state.active = false;
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
                return res;
            }

            if (wParam == VK_LEFT || wParam == VK_RIGHT || wParam == VK_UP || wParam == VK_DOWN ||
                wParam == VK_HOME || wParam == VK_END || wParam == VK_PRIOR || wParam == VK_NEXT || wParam == VK_TAB)
            {
                g_currentNumber.clear();
                if (g_state.active)
                {
                    ShowCaret(hwnd);
                    g_state.active = false;
                    UpdateResultIfPresent(hwnd, g_state.fractionIndex);
                    InvalidateRect(hwnd, nullptr, FALSE);
                }
                // VK_TAB also needs handling if we want it to move cursor normally
                if (wParam == VK_TAB)
                {
                    LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                    const LONG lenAfter = (LONG)GetWindowTextLengthW(hwnd);
                    ShiftFractionsAfter((LONG)selEnd, lenAfter - lenBefore);
                    return res;
                }
            }
            break; 
        }

        case WM_CHAR:
        {
            const wchar_t ch = (wchar_t)wParam;

            const LONG lenBefore = (LONG)GetWindowTextLengthW(hwnd);
            DWORD selStartBefore = 0, selEndBefore = 0;
            SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStartBefore, (LPARAM)&selEndBefore);

            size_t hitIndex = 0;
            if (ch == L'=')
            {
                if (IsPosInsideAnyFractionBar((LONG)selEndBefore - 1, &hitIndex))
                {
                    TriggerCalculation(hwnd, hitIndex);
                    return 0;
                }
            }

            if (!g_state.active && IsPosInsideAnyFractionBar((LONG)selEndBefore, &hitIndex))
            {
                const auto& f = g_fractions[hitIndex];
                SendMessage(hwnd, EM_SETSEL, (WPARAM)(f.barStart + f.barLen), (LPARAM)(f.barStart + f.barLen));
                SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStartBefore, (LPARAM)&selEndBefore);
            }

            if (ch == 0x0D || ch == 0x09) // Enter, Tab 
            {
                // We handle these structural keys in WM_KEYDOWN now to avoid 
                // double-processing (once in KeyDown, once in Char).
                return 0;
            }

            if (ch == 0x08) // Backspace
            {
                if (g_state.active)
                {
                    // In active mode, we DO want WM_CHAR to handle 
                    // character-by-character numerator/denominator deletion.
                    if (g_state.fractionIndex < g_fractions.size())
                    {
                        auto& f = g_fractions[g_state.fractionIndex];
                        std::wstring& target = g_state.isNumerator ? f.numerator : f.denominator;
                        if (!target.empty()) target.pop_back();

                        // Adjust bar length to fit both parts
                        const size_t maxTextLen = std::max(f.numerator.size(), f.denominator.size());
                        const LONG requiredBarLen = (LONG)std::max<size_t>(3, maxTextLen);

                        if (requiredBarLen != f.barLen)
                        {
                            const LONG delta = requiredBarLen - f.barLen;
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)(f.barStart + f.barLen));
                            std::wstring newBar((size_t)requiredBarLen, L'\u2500');
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)newBar.c_str());
                            ShiftFractionsAfter(f.barStart + f.barLen, delta);
                            f.barLen = requiredBarLen;
                        }

                        if (f.numerator.empty() && f.denominator.empty())
                        {
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)(f.barStart + f.barLen));
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                            ShiftFractionsAfter(f.barStart + 1, -f.barLen);
                            g_fractions.erase(g_fractions.begin() + (ptrdiff_t)g_state.fractionIndex);
                            ShowCaret(hwnd);
                            g_state.active = false;
                        }
                        else
                        {
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)f.barStart);
                        }

                        UpdateResultIfPresent(hwnd, g_state.fractionIndex);
                        InvalidateRect(hwnd, nullptr, FALSE);
                    }
                    return 0;
                }
                
                // When not active, backspace was already handled in WM_KEYDOWN.
                return 0;
            }

            if (ch == 27) // Escape
            {
                if (g_state.active)
                {
                    ShowCaret(hwnd);
                    g_state.active = false;
                    InvalidateRect(hwnd, nullptr, FALSE);
                    return 0;
                }
                return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            }

            if (g_state.active)
            {
                if (ch == L'=')
                {
                    ShowCaret(hwnd);
                    g_state.active = false;
                    TriggerCalculation(hwnd, g_state.fractionIndex);
                    return 0;
                }

                if (ch >= L'0' && ch <= L'9')
                {
                    if (g_state.fractionIndex < g_fractions.size())
                    {
                        auto& f = g_fractions[g_state.fractionIndex];
                        std::wstring& target = g_state.isNumerator ? f.numerator : f.denominator;
                        target.push_back(ch);

                        const size_t maxTextLen = std::max(f.numerator.size(), f.denominator.size());
                        const LONG requiredBarLen = (LONG)std::max<size_t>(3, maxTextLen);

                        if (requiredBarLen != f.barLen)
                        {
                            const LONG delta = requiredBarLen - f.barLen;
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)(f.barStart + f.barLen));
                            std::wstring newBar((size_t)requiredBarLen, L'\u2500');
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)newBar.c_str());
                            ShiftFractionsAfter(f.barStart + f.barLen, delta);
                            f.barLen = requiredBarLen;
                        }
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)f.barStart, (LPARAM)f.barStart);
                        UpdateResultIfPresent(hwnd, g_state.fractionIndex);
                        InvalidateRect(hwnd, nullptr, FALSE);
                    }
                    return 0;
                }

                // If typing anything else in active mode, drop out
                ShowCaret(hwnd);
                g_state.active = false;
                // fall through to handle key normally
            }

            if (ch == L'/')
            {
                if (!g_currentNumber.empty())
                {
                    const LONG numLen = (LONG)g_currentNumber.size();
                    const LONG numStart = (LONG)selEndBefore - numLen;
                    const LONG numEnd = (LONG)selEndBefore;

                    const LONG barLen = std::max<LONG>(3, numLen);
                    std::wstring bar((size_t)barLen, L'\u2500');
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)numStart, (LPARAM)numEnd);
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)bar.c_str());

                    ShiftFractionsAfter(numEnd, barLen - numLen);

                    FractionSpan span;
                    span.barStart = numStart;
                    span.barLen = barLen;
                    span.numerator = g_currentNumber;
                    span.denominator.clear();
                    g_fractions.push_back(std::move(span));

                    g_state.fractionIndex = g_fractions.size() - 1;
                    if (!g_state.active) HideCaret(hwnd);
                    g_state.active = true;

                    SendMessage(hwnd, EM_SETSEL, (WPARAM)(numStart + barLen), (LPARAM)(numStart + barLen));
                    InvalidateRect(hwnd, nullptr, FALSE);
                    UpdateWindow(hwnd);

                    g_currentNumber.clear();
                    return 0;
                }

                g_currentNumber.clear();
                return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            }

            if (ch >= L'0' && ch <= L'9')
                g_currentNumber += ch;
            else
                g_currentNumber.clear();

            {
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                const LONG lenAfter = (LONG)GetWindowTextLengthW(hwnd);
                ShiftFractionsAfter((LONG)selEndBefore, lenAfter - lenBefore);
                return res;
            }
        }
        }

        return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
    }

    static CHARFORMAT2W GetSelectionCharFormat(HWND hEdit)
    {
        CHARFORMAT2W cf = {};
        cf.cbSize = sizeof(cf);
        SendMessage(hEdit, EM_GETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);
        return cf;
    }

    static CHARFORMAT2W MakeNormalFormat(const CHARFORMAT2W& base)
    {
        CHARFORMAT2W cf = base;
        cf.cbSize = sizeof(cf);
        cf.dwMask = CFM_EFFECTS | CFM_OFFSET | CFM_SIZE;
        cf.dwEffects = base.dwEffects & ~(CFE_SUPERSCRIPT | CFE_SUBSCRIPT);
        cf.yOffset = 0;
        cf.yHeight = base.yHeight;
        return cf;
    }

    static CHARFORMAT2W MakeBarFormat(const CHARFORMAT2W& base)
    {
        CHARFORMAT2W cf = MakeNormalFormat(base);
        cf.dwMask |= CFM_WEIGHT | CFM_SIZE;
        cf.wWeight = FW_BOLD;
        cf.yHeight = base.yHeight;
        return cf;
    }

    static void SetInsertionFormat(HWND hEdit, const CHARFORMAT2W& cf)
    {
        SendMessage(hEdit, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);
    }
} // namespace

bool InstallFractionSupport(HWND hRichEdit)
{
    if (!hRichEdit)
        return false;

    g_hEdit = hRichEdit;
    g_currentNumber.clear();
    g_state = {};
    g_fractions.clear();

    if (!g_originalProc)
    {
        g_originalProc = (WNDPROC)SetWindowLongPtr(g_hEdit, GWLP_WNDPROC, (LONG_PTR)FractionRichEditProc);
        if (!g_originalProc)
            return false;
    }

    return true;
}

void ResetFractionSupport()
{
    g_currentNumber.clear();
    g_state = {};
    g_fractions.clear();
    if (g_hEdit)
        InvalidateRect(g_hEdit, nullptr, FALSE);
}

void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator)
{
    if (!hEdit)
        return;

    DWORD selStart = 0, selEnd = 0;
    SendMessage(hEdit, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

    const LONG startPos = (LONG)selStart;
    const LONG replacedLen = (LONG)(selEnd - selStart);

    const size_t widest = (numerator.size() > denominator.size()) ? numerator.size() : denominator.size();
    const LONG barLen = (LONG)std::max<size_t>(3, widest);
    std::wstring bar((size_t)barLen, L'\u2500');

    // Replace selection with the bar only; numerator/denominator are rendered as an overlay.
    SendMessage(hEdit, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)bar.c_str());

    FractionSpan span;
    span.barStart = startPos;
    span.barLen = barLen;
    span.numerator = numerator;
    span.denominator = denominator;
    g_fractions.push_back(std::move(span));

    // Maintain indices for any existing fractions after the insertion.
    ShiftFractionsAfter(startPos + 1, barLen - replacedLen);

    // Place caret after the bar.
    SendMessage(hEdit, EM_SETSEL, (WPARAM)(startPos + barLen), (LPARAM)(startPos + barLen));
    InvalidateRect(hEdit, nullptr, FALSE);
    UpdateWindow(hEdit);
}
