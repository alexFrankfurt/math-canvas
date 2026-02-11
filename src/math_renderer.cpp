#include "math_renderer.h"
#include "math_manager.h"
#include <richedit.h>
#include <algorithm>

// Undefine Windows min/max macros to use std::min/std::max
#ifdef max
#undef max
#endif
#ifdef min
#undef min
#endif

namespace {
    static const bool kDebugOverlay = false;

    static COLORREF GetRichEditBkColor(HWND hEdit)
    {
        COLORREF prev = (COLORREF)SendMessage(hEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)GetSysColor(COLOR_WINDOW));
        SendMessage(hEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)prev);
        return prev;
    }

    static double ComputeRenderScale(HWND hEdit, HDC hdc, const MathObject& obj, HFONT unzoomedBaseFont)
    {
        if (obj.barLen <= 0 || !unzoomedBaseFont) return 1.0;
        POINT p0 = {}, p1 = {};
        if (MathRenderer::TryGetCharPos(hEdit, obj.barStart, p0) && MathRenderer::TryGetCharPos(hEdit, obj.barStart + 1, p1))
        {
            if (p0.y == p1.y && p1.x > p0.x)
            {
                HFONT old = (HFONT)SelectObject(hdc, unzoomedBaseFont);
                wchar_t buf[16] = {0};
                TEXTRANGEW tr = { {obj.barStart, obj.barStart + 1}, buf };
                SendMessage(hEdit, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                SIZE one = {};
                GetTextExtentPoint32W(hdc, buf, 1, &one);
                SelectObject(hdc, old);
                if (one.cx > 0) return (double)(p1.x - p0.x) / (double)one.cx;
            }
        }
        return 1.0;
    }

    static HFONT CreateScaledFont(HFONT baseFont, double scale, int percent)
    {
        LOGFONTW lf = {};
        if (!baseFont || GetObjectW(baseFont, sizeof(lf), &lf) != sizeof(lf))
        {
            return CreateFontW(-11, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
                OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
        }
        if (!(scale > 0.0)) scale = 1.0;
        const int sign = (lf.lfHeight < 0) ? -1 : 1;
        const int absH = (lf.lfHeight < 0) ? -lf.lfHeight : lf.lfHeight;
        const double scaled = (double)absH * scale * ((double)percent / 100.0);
        lf.lfHeight = sign * (std::max)(1, (int)(scaled + 0.5));
        return CreateFontIndirectW(&lf);
    }
}

COLORREF MathRenderer::GetDefaultTextColor(HWND hEdit)
{
    CHARFORMAT2W cf = {};
    cf.cbSize = sizeof(cf);
    SendMessage(hEdit, EM_GETCHARFORMAT, (WPARAM)SCF_DEFAULT, (LPARAM)&cf);
    if ((cf.dwMask & CFM_COLOR) == 0 || (cf.dwEffects & CFE_AUTOCOLOR) != 0)
        return GetSysColor(COLOR_WINDOWTEXT);
    return cf.crTextColor;
}

COLORREF MathRenderer::GetActiveColor(HWND hEdit)
{
    const COLORREF text = GetDefaultTextColor(hEdit);
    const int brightness = (GetRValue(text) * 299 + GetGValue(text) * 587 + GetBValue(text) * 114) / 1000;
    return (brightness > 128) ? RGB(100, 180, 255) : RGB(0, 102, 204);
}

bool MathRenderer::TryGetCharPos(HWND hEdit, LONG charIndex, POINT& outPt)
{
    POINTL ptL = {};
    if (SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)&ptL, (LPARAM)charIndex) != -1)
    {
        outPt.x = (int)ptL.x; outPt.y = (int)ptL.y;
        return true;
    }
    LRESULT xy = SendMessage(hEdit, EM_POSFROMCHAR, (WPARAM)charIndex, 0);
    if (xy != -1)
    {
        outPt.x = (short)LOWORD(xy); outPt.y = (short)HIWORD(xy);
        return true;
    }
    return false;
}

void MathRenderer::Draw(HWND hEdit, HDC hdc, const MathObject& obj, size_t objIndex, const MathTypingState& state)
{
    if (obj.barLen <= 0) return;
    POINT ptStart = {}, ptEnd = {};
    if (!TryGetCharPos(hEdit, obj.barStart, ptStart)) return;
    if (!TryGetCharPos(hEdit, obj.barStart + std::max<LONG>(0, obj.barLen - 1), ptEnd)) return;

    HFONT baseFont = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
    if (!baseFont) baseFont = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

    double renderScale = ComputeRenderScale(hEdit, hdc, obj, baseFont);
    // Non-fraction command objects use 2x-height anchors for vertical space,
    // but the renderScale picks up that 2x width too. Compensate so drawn
    // content stays at the correct size relative to zoom.
    if (obj.type == MathType::Summation || obj.type == MathType::Integral
        || obj.type == MathType::SystemOfEquations)
        renderScale *= 0.5;
    HFONT renderBaseFont = CreateScaledFont(baseFont, renderScale, 100);
    HFONT limitFont = CreateScaledFont(baseFont, renderScale, 70);

    const int saved = SaveDC(hdc);
    // Clip to the RichEdit's client rect so we never draw over the toolbar
    // or outside the control's visible area.
    RECT clientRect;
    GetClientRect(hEdit, &clientRect);
    HRGN hClip = CreateRectRgnIndirect(&clientRect);
    SelectClipRgn(hdc, hClip);
    DeleteObject(hClip);
    SetBkMode(hdc, TRANSPARENT);
    SetTextAlign(hdc, TA_BASELINE | TA_CENTER);
    HFONT oldFont = (HFONT)SelectObject(hdc, renderBaseFont);
    TEXTMETRICW tmBase = {};
    GetTextMetricsW(hdc, &tmBase);

    const int barWidth = (ptEnd.x - ptStart.x) + (int)(tmBase.tmAveCharWidth);
    const int xCenter = ptStart.x + (barWidth / 2);
    const int yMid = ptStart.y + tmBase.tmAscent;
    const COLORREF normalColor = GetDefaultTextColor(hEdit);
    const COLORREF activeColor = GetActiveColor(hEdit);

    auto DrawPart = [&](const std::wstring& text, int x, int y, int partIdx) {
        bool isActive = (state.active && state.objectIndex == objIndex && state.activePart == partIdx);
        SetTextColor(hdc, isActive ? activeColor : normalColor);
        
        if (kDebugOverlay)
        {
            SIZE sz = {};
            GetTextExtentPoint32W(hdc, text.empty() ? L"?" : text.c_str(), (int)std::max<size_t>(1, text.size()), &sz);
            RECT rc = { x - sz.cx/2, y - sz.cy, x + sz.cx/2, y };
            if (GetTextAlign(hdc) & TA_LEFT) {
                rc.left = x; rc.right = x + sz.cx;
            }
            HBRUSH hBr = CreateSolidBrush(isActive ? RGB(255, 0, 0) : RGB(0, 255, 0));
            FrameRect(hdc, &rc, hBr);
            DeleteObject(hBr);
        }

        if (text.empty() && isActive) TextOutW(hdc, x, y, L"?", 1);
        else TextOutW(hdc, x, y, text.c_str(), (int)text.size());
    };

    if (kDebugOverlay)
    {
        HPEN hPen = CreatePen(PS_DOT, 1, RGB(200, 200, 200));
        HPEN oldP = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, ptStart.x - 20, yMid, NULL);
        LineTo(hdc, ptEnd.x + 40, yMid);
        MoveToEx(hdc, xCenter, yMid - 40, NULL);
        LineTo(hdc, xCenter, yMid + 40);
        SelectObject(hdc, oldP);
        DeleteObject(hPen);
    }

    // (Anchor characters are hidden via CHARFORMAT in math_editor.cpp,
    //  so we only need to paint-over for non-Fraction types that may
    //  still show raw anchor glyphs.)
    if (obj.type != MathType::Fraction)
    {
        COLORREF bk = GetRichEditBkColor(hEdit);
        SIZE charSize = {};
        GetTextExtentPoint32W(hdc, L"\u2500", 1, &charSize);
        int coverRight = ptEnd.x + (std::max)((int)charSize.cx, (int)tmBase.tmAveCharWidth) + 4;
        RECT rc = { ptStart.x - 2, ptStart.y - 4, coverRight, ptStart.y + tmBase.tmHeight + 4 };
        HBRUSH hBr = CreateSolidBrush(bk);
        FillRect(hdc, &rc, hBr);
        DeleteObject(hBr);
    }

    if (obj.type == MathType::Fraction)
    {
        // 50% bigger than the default limit font (70% * 1.5 ≈ 105%)
        HFONT fracFont = CreateScaledFont(baseFont, renderScale, 105);
        SelectObject(hdc, fracFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);

        const int gap = 4;  // pixels between bar and text

        // Place the bar at the vertical center of the anchor cell.
        const int barY = ptStart.y + tmBase.tmHeight / 2;

        // Numerator: baseline so bottom of text is `gap` above the bar
        DrawPart(obj.part1, xCenter, barY - gap - tmL.tmDescent, 1);
        // Denominator: baseline so top of text is `gap` below the bar
        DrawPart(obj.part2, xCenter, barY + gap + tmL.tmAscent, 2);

        SelectObject(hdc, oldFont);
        DeleteObject(fracFont);

        // Draw the vinculum (fraction bar) via GDI
        HPEN hPen = CreatePen(PS_SOLID, (std::max)(1, (int)(1.2 * renderScale)), normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, ptStart.x, barY, NULL);
        LineTo(hdc, ptStart.x + barWidth, barY);
        SelectObject(hdc, oldPen);
        DeleteObject(hPen);

        // Draw result to the right, vertically centered on the bar
        if (!obj.resultText.empty()) {
            SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TEXTMETRICW tmR = {}; GetTextMetricsW(hdc, &tmR);
            // Align equals sign's horizontal stroke with the bar
            int resultBaseline = barY + (tmR.tmAscent - tmR.tmDescent) / 2;
            SetTextColor(hdc, activeColor);
            TextOutW(hdc, ptStart.x + barWidth + 4, resultBaseline, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }
    }
    else if (obj.type == MathType::Summation)
    {
        // Draw sigma symbol via GDI in Cambria Math
        {
            LOGFONTW lfSym = {};
            GetObjectW(renderBaseFont, sizeof(lfSym), &lfSym);
            wcscpy_s(lfSym.lfFaceName, L"Cambria Math");
            HFONT symbolFont = CreateFontIndirectW(&lfSym);
            HFONT prevF = (HFONT)SelectObject(hdc, symbolFont);
            SetTextColor(hdc, normalColor);
            TextOutW(hdc, xCenter, yMid, L"\u2211", 1);
            SelectObject(hdc, prevF);
            DeleteObject(symbolFont);
        }

        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);
        
        // Upper limit: center horizontally at xCenter.
        // Y: yMid - tmBase.tmAscent (top of sigma) - 2
        DrawPart(obj.part1, xCenter, yMid - tmBase.tmAscent - 2, 1);
        
        // Lower limit: center horizontally at xCenter.
        // Y: yMid + tmBase.tmDescent + tmL.tmAscent + 2
        DrawPart(obj.part2, xCenter, yMid + tmBase.tmDescent + tmL.tmAscent + 2, 2);
        
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        // Draw expression slightly to the right of the anchor spaces
        DrawPart(obj.part3, ptEnd.x + 4, yMid, 3);

        // Draw result (e.g. " ＝ 302") right after expression via GDI
        if (!obj.resultText.empty()) {
            SIZE exprSz = {};
            GetTextExtentPoint32W(hdc, obj.part3.c_str(), (int)obj.part3.size(), &exprSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, ptEnd.x + 4 + exprSz.cx + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }
    }
    else if (obj.type == MathType::Integral)
    {
        // Draw integral symbol via GDI in Cambria Math
        {
            LOGFONTW lfSym = {};
            GetObjectW(renderBaseFont, sizeof(lfSym), &lfSym);
            wcscpy_s(lfSym.lfFaceName, L"Cambria Math");
            HFONT symbolFont = CreateFontIndirectW(&lfSym);
            HFONT prevF = (HFONT)SelectObject(hdc, symbolFont);
            SetTextColor(hdc, normalColor);
            TextOutW(hdc, xCenter, yMid, L"\u222B", 1);
            SelectObject(hdc, prevF);
            DeleteObject(symbolFont);
        }

        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, limitFont);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);

        // Upper limit: side-aligned, slightly right.
        DrawPart(obj.part1, ptEnd.x - 2, yMid - tmBase.tmAscent + (int)(tmBase.tmAscent * 0.2), 1);
        // Lower limit: side-aligned, slightly left.
        DrawPart(obj.part2, ptEnd.x - 8, yMid + tmBase.tmDescent + 2, 2);

        SelectObject(hdc, renderBaseFont);
        DrawPart(obj.part3, ptEnd.x + 6, yMid, 3);

        // Draw result right after expression via GDI
        if (!obj.resultText.empty()) {
            SIZE exprSz = {};
            GetTextExtentPoint32W(hdc, obj.part3.c_str(), (int)obj.part3.size(), &exprSz);
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, ptEnd.x + 6 + exprSz.cx + 4, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }
    }
    else if (obj.type == MathType::SystemOfEquations)
    {
        // First measure the equations block to know total height
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmEq = {};
        GetTextMetricsW(hdc, &tmEq);
        int lineH = tmEq.tmHeight + 4;
        int eqCount = 3;
        if (obj.part3.empty() && !(state.active && state.objectIndex == objIndex && state.activePart == 3))
            eqCount = 2;
        int totalH = lineH * eqCount;
        // yTop is the baseline of the first equation line
        int yTop = yMid - totalH / 2 + tmEq.tmAscent;

        // Draw a left curly brace using GDI Bezier curves for precise height control
        {
            int blockTop = yTop - tmEq.tmAscent - 2;
            int blockBot = blockTop + totalH + 4;
            int blockMidY = (blockTop + blockBot) / 2;
            int braceW = (int)(tmBase.tmAveCharWidth * 1.2);
            if (braceW < 10) braceW = 10;
            int xRight = ptStart.x + braceW;           // where the arms end (top & bottom)
            int xMid   = ptStart.x + braceW / 2;       // vertical spine of the brace
            int xTip   = ptStart.x;                     // middle tip pointing left
            int armH   = (blockMidY - blockTop);        // half-height

            HPEN hPen = CreatePen(PS_SOLID, (std::max)(1, (int)(1.5 * renderScale)), normalColor);
            HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
            HBRUSH oldBrush = (HBRUSH)SelectObject(hdc, GetStockObject(NULL_BRUSH));

            // Top hook: small curve from top endpoint inward
            // Goes from (xRight, blockTop) curving left to (xMid, blockTop + armH*0.25)
            POINT topHook[4] = {
                { xRight, blockTop },
                { xMid,   blockTop },
                { xMid,   blockTop },
                { xMid,   blockTop + (int)(armH * 0.25) }
            };
            PolyBezier(hdc, topHook, 4);

            // Top arm: straight-ish from spine down to the middle tip
            // Goes from (xMid, blockTop + armH*0.25) curving to (xTip, blockMidY)
            POINT topArm[4] = {
                { xMid, blockTop + (int)(armH * 0.25) },
                { xMid, blockMidY - (int)(armH * 0.15) },
                { xMid, blockMidY - (int)(armH * 0.05) },
                { xTip, blockMidY }
            };
            PolyBezier(hdc, topArm, 4);

            // Bottom arm: from middle tip curving back to spine
            POINT botArm[4] = {
                { xTip, blockMidY },
                { xMid, blockMidY + (int)(armH * 0.05) },
                { xMid, blockMidY + (int)(armH * 0.15) },
                { xMid, blockBot - (int)(armH * 0.25) }
            };
            PolyBezier(hdc, botArm, 4);

            // Bottom hook: from spine curving out to bottom endpoint
            POINT botHook[4] = {
                { xMid, blockBot - (int)(armH * 0.25) },
                { xMid, blockBot },
                { xMid, blockBot },
                { xRight, blockBot }
            };
            PolyBezier(hdc, botHook, 4);

            SelectObject(hdc, oldBrush);
            SelectObject(hdc, oldPen);
            DeleteObject(hPen);

            // Draw equations stacked vertically, left-aligned after the brace
            int eqX = ptStart.x + braceW + 6;
            SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
            SelectObject(hdc, renderBaseFont);
            DrawPart(obj.part1, eqX, yTop, 1);
            DrawPart(obj.part2, eqX, yTop + lineH, 2);
            if (eqCount >= 3)
                DrawPart(obj.part3, eqX, yTop + lineH * 2, 3);

            // Draw result to the right, vertically centered on the brace
            if (!obj.resultText.empty()) {
                SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
                HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
                LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
                lfB.lfWeight = FW_BOLD;
                DeleteObject(boldFont);
                boldFont = CreateFontIndirectW(&lfB);
                HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
                TEXTMETRICW tmR = {}; GetTextMetricsW(hdc, &tmR);
                
                // Find the rightmost extent of the equations
                SIZE eq1Sz = {}, eq2Sz = {}, eq3Sz = {};
                GetTextExtentPoint32W(hdc, obj.part1.c_str(), (int)obj.part1.size(), &eq1Sz);
                GetTextExtentPoint32W(hdc, obj.part2.c_str(), (int)obj.part2.size(), &eq2Sz);
                if (eqCount >= 3) GetTextExtentPoint32W(hdc, obj.part3.c_str(), (int)obj.part3.size(), &eq3Sz);
                
                int maxEqWidth = std::max(std::max(eq1Sz.cx, eq2Sz.cx), eq3Sz.cx);
                int resultX = eqX + maxEqWidth + 10; // Add some padding
                int resultBaseline = yMid; // Vertically center with the brace
                
                SetTextColor(hdc, activeColor);
                TextOutW(hdc, resultX, resultBaseline, obj.resultText.c_str(), (int)obj.resultText.size());
                SelectObject(hdc, prevF);
                DeleteObject(boldFont);
            }
        }
    }
    else if (obj.type == MathType::SquareRoot)
    {
        // Measure the expression text to determine radical size
        SelectObject(hdc, renderBaseFont);
        TEXTMETRICW tmExpr = {};
        GetTextMetricsW(hdc, &tmExpr);

        // Get expression text size (part1 is the radicand)
        SIZE exprSz = {};
        const wchar_t* exprText = obj.part1.empty() ? L"?" : obj.part1.c_str();
        int exprLen = obj.part1.empty() ? 1 : (int)obj.part1.size();
        GetTextExtentPoint32W(hdc, exprText, exprLen, &exprSz);

        // Layout constants
        const int pad = (std::max)(2, (int)(tmBase.tmHeight / 8));
        const int overlineGap = (std::max)(2, (int)(tmBase.tmHeight / 10));
        const int radicalW = (int)(tmBase.tmAveCharWidth * 1.0);  // width of the radical sign part
        const int penWidth = (std::max)(1, (int)(1.2 * renderScale));

        // Radical vertical extents
        int radTop = yMid - tmExpr.tmAscent - overlineGap - penWidth;  // top of the overline
        int radBot = yMid + tmExpr.tmDescent + pad;                    // bottom of radical
        int radMid = radBot - (int)((radBot - radTop) * 0.35);         // the "valley" of the checkmark

        // X positions
        int xRadStart = ptStart.x;                        // leftmost point (short leading stroke)
        int xValley   = xRadStart + radicalW / 3;         // bottom of V
        int xRadPeak  = xRadStart + radicalW;              // top of the upstroke, starts overline
        int xExprStart = xRadPeak + pad;                   // where expression text begins
        int xOverlineEnd = xExprStart + exprSz.cx + pad;   // right end of overline

        // Draw the radical sign using polyline
        HPEN hPen = CreatePen(PS_SOLID, penWidth, normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);

        // Leading short horizontal-ish stroke
        MoveToEx(hdc, xRadStart, radMid - (int)(tmBase.tmHeight * 0.05), NULL);
        LineTo(hdc, xValley, radBot);       // down to the valley
        LineTo(hdc, xRadPeak, radTop);      // up to the peak
        LineTo(hdc, xOverlineEnd, radTop);  // overline across the top

        // Small tail on the right end of overline going down slightly
        LineTo(hdc, xOverlineEnd, radTop + (int)(tmBase.tmHeight * 0.1));

        SelectObject(hdc, oldPen);
        DeleteObject(hPen);

        // Draw the expression (part1) under the overline
        SetTextAlign(hdc, TA_BASELINE | TA_LEFT);
        SelectObject(hdc, renderBaseFont);
        DrawPart(obj.part1, xExprStart, yMid, 1);

        // Draw result right after the radical if present
        if (!obj.resultText.empty()) {
            SetTextColor(hdc, activeColor);
            HFONT boldFont = CreateScaledFont(baseFont, renderScale, 100);
            LOGFONTW lfB = {}; GetObjectW(boldFont, sizeof(lfB), &lfB);
            lfB.lfWeight = FW_BOLD;
            DeleteObject(boldFont);
            boldFont = CreateFontIndirectW(&lfB);
            HFONT prevF = (HFONT)SelectObject(hdc, boldFont);
            TextOutW(hdc, xOverlineEnd + 6, yMid, obj.resultText.c_str(), (int)obj.resultText.size());
            SelectObject(hdc, prevF);
            DeleteObject(boldFont);
        }
    }
    RestoreDC(hdc, saved);
    DeleteObject(renderBaseFont); DeleteObject(limitFont);
}

bool MathRenderer::GetHitPart(HWND hEdit, HDC hdc, POINT ptMouse, size_t* outIndex, int* outPart)
{
    auto& objects = MathManager::Get().GetObjects();
    for (size_t i = 0; i < objects.size(); ++i)
    {
        const auto& obj = objects[i];
        POINT ptS = {}, ptE = {};
        if (!TryGetCharPos(hEdit, obj.barStart, ptS)) continue;
        if (!TryGetCharPos(hEdit, obj.barStart + std::max<LONG>(0, obj.barLen - 1), ptE)) continue;

        HFONT baseF = (HFONT)SendMessage(hEdit, WM_GETFONT, 0, 0);
        if (!baseF) baseF = (HFONT)GetStockObject(DEFAULT_GUI_FONT);

        const double scale = ComputeRenderScale(hEdit, hdc, obj, baseF);
        HFONT limitF = CreateScaledFont(baseF, scale, 70);
        HFONT baseRF = CreateScaledFont(baseF, scale, 100);

        TEXTMETRICW tmB = {};
        HFONT old = (HFONT)SelectObject(hdc, baseRF);
        GetTextMetricsW(hdc, &tmB);

        const int bW = (ptE.x - ptS.x) + (int)(tmB.tmAveCharWidth * scale);
        const int xC = ptS.x + (bW / 2);
        const int yM = ptS.y + tmB.tmAscent;
        const int gap = std::max<int>(2, (int)(tmB.tmHeight / 10));

        bool hit = false;
        auto Check = [&](RECT rc, int pIdx) {
            if (PtInRect(&rc, ptMouse)) { *outIndex = i; *outPart = pIdx; return true; }
            return false;
        };

        if (obj.type == MathType::Fraction)
        {
            SelectObject(hdc, limitF); SIZE sz = {};
            GetTextExtentPoint32W(hdc, obj.part1.c_str(), (int)obj.part1.size(), &sz);
            if (Check({ xC - sz.cx/2 - 10, yM - (int)(tmB.tmAscent*0.4) - sz.cy - 10, xC + sz.cx/2 + 10, yM - (int)(tmB.tmAscent*0.4) + 5 }, 1)) hit = true;
            if (!hit) {
                GetTextExtentPoint32W(hdc, obj.part2.c_str(), (int)obj.part2.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM + (int)(tmB.tmDescent*0.4) - 5, xC + sz.cx/2 + 10, yM + (int)(tmB.tmDescent*0.4) + sz.cy + 10 }, 2)) hit = true;
            }
        }
        else if (obj.type == MathType::Summation)
        {
            SelectObject(hdc, limitF); SIZE sz = {};
            GetTextExtentPoint32W(hdc, obj.part1.c_str(), (int)obj.part1.size(), &sz);
            if (Check({ xC - sz.cx/2 - 10, yM - tmB.tmAscent - sz.cy - 10, xC + sz.cx/2 + 10, yM - tmB.tmAscent + 5 }, 1)) hit = true;
            if (!hit) {
                GetTextExtentPoint32W(hdc, obj.part2.c_str(), (int)obj.part2.size(), &sz);
                if (Check({ xC - sz.cx/2 - 10, yM + tmB.tmDescent - 5, xC + sz.cx/2 + 10, yM + tmB.tmDescent + sz.cy + 10 }, 2)) hit = true;
            }
            if (!hit) {
                SelectObject(hdc, baseRF);
                GetTextExtentPoint32W(hdc, obj.part3.empty() ? L"?" : obj.part3.c_str(), (int)std::max<size_t>(1, obj.part3.size()), &sz);
                if (Check({ ptE.x + 2, yM - tmB.tmAscent, ptE.x + sz.cx + 20, yM + tmB.tmDescent }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::Integral)
        {
            SelectObject(hdc, limitF); SIZE sz = {};
            GetTextExtentPoint32W(hdc, obj.part1.c_str(), (int)obj.part1.size(), &sz);
            if (Check({ ptE.x - 2, yM - tmB.tmAscent - 5, ptE.x + sz.cx + 10, yM - tmB.tmAscent + sz.cy + 5 }, 1)) hit = true;
            if (!hit) {
                GetTextExtentPoint32W(hdc, obj.part2.c_str(), (int)obj.part2.size(), &sz);
                if (Check({ ptE.x - 10, yM + tmB.tmDescent - 5, ptE.x + sz.cx + 5, yM + tmB.tmDescent + sz.cy + 5 }, 2)) hit = true;
            }
            if (!hit) {
                SelectObject(hdc, baseRF);
                GetTextExtentPoint32W(hdc, obj.part3.empty() ? L"?" : obj.part3.c_str(), (int)std::max<size_t>(1, obj.part3.size()), &sz);
                if (Check({ ptE.x + 4, yM - tmB.tmAscent, ptE.x + sz.cx + 20, yM + tmB.tmDescent }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::SystemOfEquations)
        {
            // Compute brace width matching the renderer
            int braceW = (int)(tmB.tmAveCharWidth * 1.2);
            if (braceW < 10) braceW = 10;

            SelectObject(hdc, baseRF);
            TEXTMETRICW tmEq = {};
            GetTextMetricsW(hdc, &tmEq);
            int lineH = tmEq.tmHeight + 4;
            int eqCount = 3;
            if (obj.part3.empty()) eqCount = 2;
            int totalH = lineH * eqCount;
            int yTop = yM - totalH / 2;
            int eqX = ptS.x + braceW + 6;

            SIZE sz = {};
            // Eq 1 hit area
            GetTextExtentPoint32W(hdc, obj.part1.empty() ? L"?" : obj.part1.c_str(), (int)std::max<size_t>(1, obj.part1.size()), &sz);
            if (Check({ eqX - 5, yTop, eqX + (std::max)(sz.cx, (LONG)40) + 10, yTop + lineH }, 1)) hit = true;
            // Eq 2 hit area
            if (!hit) {
                GetTextExtentPoint32W(hdc, obj.part2.empty() ? L"?" : obj.part2.c_str(), (int)std::max<size_t>(1, obj.part2.size()), &sz);
                if (Check({ eqX - 5, yTop + lineH, eqX + (std::max)(sz.cx, (LONG)40) + 10, yTop + lineH * 2 }, 2)) hit = true;
            }
            // Eq 3 hit area
            if (!hit && eqCount >= 3) {
                GetTextExtentPoint32W(hdc, obj.part3.empty() ? L"?" : obj.part3.c_str(), (int)std::max<size_t>(1, obj.part3.size()), &sz);
                if (Check({ eqX - 5, yTop + lineH * 2, eqX + (std::max)(sz.cx, (LONG)40) + 10, yTop + lineH * 3 }, 3)) hit = true;
            }
        }
        else if (obj.type == MathType::SquareRoot)
        {
            SelectObject(hdc, baseRF);
            TEXTMETRICW tmExpr = {};
            GetTextMetricsW(hdc, &tmExpr);

            SIZE sz = {};
            GetTextExtentPoint32W(hdc, obj.part1.empty() ? L"?" : obj.part1.c_str(), (int)std::max<size_t>(1, obj.part1.size()), &sz);

            int pad = std::max<int>(2, (int)(tmB.tmHeight / 8));
            int radicalW = (int)(tmB.tmAveCharWidth * 1.0);
            int xExprStart = ptS.x + radicalW + pad;
            int overlineGap = std::max<int>(2, (int)(tmB.tmHeight / 10));
            int radTop = yM - tmExpr.tmAscent - overlineGap - 2;

            // Hit area covers the entire radical expression area
            RECT rcExpr = { ptS.x, radTop - 5, xExprStart + sz.cx + pad + 10, yM + tmExpr.tmDescent + pad + 5 };
            if (Check(rcExpr, 1)) hit = true;
        }
        SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF);
        if (hit) return true;
    }
    return false;
}
