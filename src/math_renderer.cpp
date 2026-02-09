#include "math_renderer.h"
#include "math_manager.h"
#include <richedit.h>
#include <algorithm>

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

    const double renderScale = ComputeRenderScale(hEdit, hdc, obj, baseFont);
    HFONT renderBaseFont = CreateScaledFont(baseFont, renderScale, 100);
    HFONT limitFont = CreateScaledFont(baseFont, renderScale, 70);

    const int saved = SaveDC(hdc);
    SelectClipRgn(hdc, nullptr);
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

    // For Summation/Integral: paint over the double-height NBSP anchor chars
    // to hide rendering artifacts from the RichEdit control
    if (obj.type != MathType::Fraction)
    {
        COLORREF bk = GetRichEditBkColor(hEdit);
        RECT rc = { ptStart.x, ptStart.y, ptStart.x + barWidth, ptStart.y + tmBase.tmHeight };
        HBRUSH hBr = CreateSolidBrush(bk);
        FillRect(hdc, &rc, hBr);
        DeleteObject(hBr);
    }

    if (obj.type == MathType::Fraction)
    {
        SelectObject(hdc, limitFont);
        DrawPart(obj.part1, xCenter, yMid - (int)(tmBase.tmAscent * 0.4) - 2, 1);
        TEXTMETRICW tmL = {}; GetTextMetricsW(hdc, &tmL);
        DrawPart(obj.part2, xCenter, yMid + (int)(tmBase.tmDescent * 0.4) + tmL.tmAscent + 2, 2);
        SelectObject(hdc, oldFont);
        HPEN hPen = CreatePen(PS_SOLID, (int)(1 * renderScale), normalColor);
        HPEN oldPen = (HPEN)SelectObject(hdc, hPen);
        MoveToEx(hdc, ptStart.x, yMid - (int)(tmBase.tmAscent * 0.1), NULL);
        LineTo(hdc, ptEnd.x + (int)(tmBase.tmAveCharWidth), yMid - (int)(tmBase.tmAscent * 0.1));
        SelectObject(hdc, oldPen); DeleteObject(hPen);
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
        SelectObject(hdc, old); DeleteObject(baseRF); DeleteObject(limitF);
        if (hit) return true;
    }
    return false;
}
