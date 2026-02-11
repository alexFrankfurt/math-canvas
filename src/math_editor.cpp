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

    // Structure to hold result of equation solving with detailed error messages
    struct EquationResult {
        double value;
        std::wstring message;
        bool success;

        EquationResult() : value(0), success(false) {}
        EquationResult(double v, const std::wstring& msg, bool s) : value(v), message(msg), success(s) {}
    };

namespace
{
    HWND g_hEdit = nullptr;
    WNDPROC g_originalProc = nullptr;
    std::wstring g_currentCommand;
    std::wstring g_currentNumber;
    bool g_suppressNextChar = false;

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
            // For other math types (fractions, etc.), use the generic calculation
            double result = mgr.CalculateResult(obj);
            wchar_t resultBuf[128];
            if (result == (long long)result) swprintf_s(resultBuf, L" %lld", (long long)result);
            else swprintf_s(resultBuf, L" %.3g", result);
            obj.resultText = std::wstring(L" \uFF1D") + resultBuf;
        }
        InvalidateRect(hwnd, nullptr, TRUE);
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
            InvalidateRect(hwnd, nullptr, TRUE);
            return;
        }

        // Handle other math types (fractions, etc.)
        double result = mgr.CalculateResult(obj);
        wchar_t resultBuf[128];
        if (result == (long long)result) swprintf_s(resultBuf, L" %lld", (long long)result);
        else swprintf_s(resultBuf, L" %.3g", result);

        obj.resultText = std::wstring(L" \uFF1D") + resultBuf;
        SendMessage(hwnd, EM_SETSEL, obj.barStart + obj.barLen, obj.barStart + obj.barLen);
        InvalidateRect(hwnd, nullptr, TRUE);
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

        case WM_LBUTTONDOWN:
        {
            POINT pt = { (short)LOWORD(lParam), (short)HIWORD(lParam) };
            HDC hdc = GetDC(hwnd); size_t idx = 0; int part = 0;
            bool hit = MathRenderer::GetHitPart(hwnd, hdc, pt, &idx, &part);
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
                SendMessage(hwnd, EM_SETSEL, (WPARAM)objects[idx].barStart, (LPARAM)objects[idx].barStart);
                InvalidateRect(hwnd, nullptr, TRUE);
                return 0; 
            }
            else
            {
                if (state.active) { ShowCaret(hwnd); state.active = false; InvalidateRect(hwnd, nullptr, TRUE); }
            }
            break; 
        }

        case WM_PAINT:
        case WM_PRINTCLIENT:
        {
            const LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            HDC hdc = (uMsg == WM_PAINT) ? GetDC(hwnd) : (HDC)wParam;
            if (hdc)
            {
                if (state.active) HideCaret(hwnd);
                for (size_t i = 0; i < objects.size(); ++i)
                    MathRenderer::Draw(hwnd, hdc, objects[i], i, state);
                if (uMsg == WM_PAINT) ReleaseDC(hwnd, hdc);
            }
            return res;
        }

        case WM_MOUSEWHEEL:
        case WM_VSCROLL:
        case WM_HSCROLL:
        {
            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            InvalidateRect(hwnd, nullptr, TRUE);
            return res;
        }

        case WM_KEYDOWN:
        {
            const LONG lenBefore = (LONG)GetWindowTextLengthW(hwnd);
            DWORD selStart = 0, selEnd = 0;
            SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

            if (wParam == VK_BACK || wParam == VK_DELETE) {
                if (selStart != selEnd) mgr.DeleteObjectsInRange((LONG)selStart, (LONG)selEnd);
            }

            if (wParam == VK_RETURN)
            {
                if (state.active) {
                    auto& obj = objects[state.objectIndex];
                    LONG afterObj = obj.barStart + obj.barLen;
                    state.active = false; 
                    

                    // For system of equations, trigger calculation on Enter regardless of position
                    if (obj.type == MathType::SystemOfEquations) {
                        // Simple validation - just make sure both parts exist
                        if (obj.part1.empty() || obj.part2.empty()) {
                            return 0;
                        }

                        // Proceed with calculation
                        TriggerCalculation(hwnd, state.objectIndex);
                        InvalidateRect(hwnd, nullptr, TRUE);  // Force immediate UI refresh
                    }
                    UpdateResultIfPresent(hwnd, state.objectIndex);
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)afterObj, (LPARAM)afterObj);
                    ShowCaret(hwnd); InvalidateRect(hwnd, nullptr, TRUE);
                    g_suppressNextChar = true;
                    return 0; 
                }
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                mgr.ShiftObjectsAfter((LONG)selEnd, (LONG)GetWindowTextLengthW(hwnd) - lenBefore);
                if (!objects.empty()) InvalidateRect(hwnd, nullptr, TRUE);
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
                    InvalidateRect(hwnd, nullptr, TRUE);
                    return 0;
                }
            }

            if (wParam == VK_LEFT || wParam == VK_RIGHT || wParam == VK_UP || wParam == VK_DOWN ||
                wParam == VK_HOME || wParam == VK_END || wParam == VK_PRIOR || wParam == VK_NEXT || wParam == VK_TAB)
            {
                g_currentNumber.clear(); g_currentCommand.clear();
                if (state.active) {
                    if (wParam == VK_TAB) {
                        int maxP = (objects[state.objectIndex].type == MathType::Fraction) ? 2 : 3;
                        if (objects[state.objectIndex].type == MathType::SystemOfEquations) maxP = 3;
                        if (objects[state.objectIndex].type == MathType::SquareRoot) maxP = 1;
                        state.activePart = (state.activePart % maxP) + 1;
                        InvalidateRect(hwnd, nullptr, TRUE);
                        return 0;
                    }
                    auto& obj = objects[state.objectIndex];
                    LONG afterObj = obj.barStart + obj.barLen;
                    ShowCaret(hwnd); state.active = false;
                    UpdateResultIfPresent(hwnd, state.objectIndex);
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)afterObj, (LPARAM)afterObj);
                    InvalidateRect(hwnd, nullptr, TRUE);
                }
            }
            {
                LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
                mgr.ShiftObjectsAfter((LONG)selEnd, (LONG)GetWindowTextLengthW(hwnd) - lenBefore);
                if (!objects.empty()) InvalidateRect(hwnd, nullptr, TRUE);
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
                    std::wstring& target = (state.activePart == 1) ? obj.part1 : (state.activePart == 2 ? obj.part2 : obj.part3);
                    if (!target.empty()) target.pop_back();

                    if (obj.type == MathType::Fraction) {
                        const LONG reqLen = (LONG)std::max<size_t>(3, std::max(obj.part1.size(), obj.part2.size()));
                        if (reqLen != obj.barLen) {
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + obj.barLen));
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)reqLen, L'\u2500').c_str());
                            HideAnchorChars(hwnd, obj.barStart, reqLen);
                            mgr.ShiftObjectsAfter(obj.barStart + obj.barLen, reqLen - obj.barLen);
                            obj.barLen = reqLen;
                        }
                    }

                    if (obj.part1.empty() && obj.part2.empty() && obj.type == MathType::Fraction) {
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + obj.barLen));
                        SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)L"");
                        mgr.ShiftObjectsAfter(obj.barStart + 1, -obj.barLen);
                        objects.erase(objects.begin() + state.objectIndex);
                        ShowCaret(hwnd); state.active = false;
                    } else {
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                    }
                    UpdateResultIfPresent(hwnd, state.objectIndex);
                    InvalidateRect(hwnd, nullptr, TRUE);
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
                        // For other types (fractions, etc.), trigger calculation normally
                        ShowCaret(hwnd); 
                        state.active = false; 
                        TriggerCalculation(hwnd, state.objectIndex); 
                        return 0; 
                    }
                    // For system of equations, allow the equals sign to be typed normally
                    if (state.objectIndex < objects.size()) {
                        auto& obj = objects[state.objectIndex];
                        std::wstring& target = (state.activePart == 1) ? obj.part1 : (state.activePart == 2 ? obj.part2 : obj.part3);
                        target.push_back(ch);

                        if (obj.type == MathType::Fraction) {
                            const LONG reqLen = (LONG)std::max<size_t>(3, std::max(obj.part1.size(), obj.part2.size()));
                            if (reqLen != obj.barLen) {
                                SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + obj.barLen));
                                SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)reqLen, L'\u2500').c_str());
                                HideAnchorChars(hwnd, obj.barStart, reqLen);
                                mgr.ShiftObjectsAfter(obj.barStart + obj.barLen, reqLen - obj.barLen);
                                obj.barLen = reqLen;
                            }
                        }
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                        UpdateResultIfPresent(hwnd, state.objectIndex);
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                    return 0;
                }
                if (iswprint(ch) && ch != L'^' && ch != L'_') {
                    if (state.objectIndex < objects.size()) {
                        auto& obj = objects[state.objectIndex];
                        std::wstring& target = (state.activePart == 1) ? obj.part1 : (state.activePart == 2 ? obj.part2 : obj.part3);
                        if (state.activePart == 3 && target == L"{}") target = std::wstring(1, L'{') + ch + L'}';
                        else if (state.activePart == 3 && target.size() >= 2 && target.front() == L'{' && target.back() == L'}') target.insert(target.size() - 1, 1, ch);
                        else target.push_back(ch);

                        if (obj.type == MathType::Fraction) {
                            const LONG reqLen = (LONG)std::max<size_t>(3, std::max(obj.part1.size(), obj.part2.size()));
                            if (reqLen != obj.barLen) {
                                SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)(obj.barStart + obj.barLen));
                                SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)reqLen, L'\u2500').c_str());
                                HideAnchorChars(hwnd, obj.barStart, reqLen);
                                mgr.ShiftObjectsAfter(obj.barStart + obj.barLen, reqLen - obj.barLen);
                                obj.barLen = reqLen;
                            }
                        }
                        SendMessage(hwnd, EM_SETSEL, (WPARAM)obj.barStart, (LPARAM)obj.barStart);
                        UpdateResultIfPresent(hwnd, state.objectIndex);
                        InvalidateRect(hwnd, nullptr, TRUE);
                    }
                    return 0;
                }
                if (ch == L'^') { state.activePart = 1; InvalidateRect(hwnd, nullptr, TRUE); return 0; }
                if (ch == L'_') { state.activePart = 2; InvalidateRect(hwnd, nullptr, TRUE); return 0; }
                ShowCaret(hwnd); state.active = false;
            }

            if ((ch == L' ' || ch == L'_' || ch == L'^') && !g_currentCommand.empty())
            {
                MathObject obj; bool found = false;
                wchar_t anchorStr[10] = { 0 };
                if (g_currentCommand == L"\\sum") {
                    obj.type = MathType::Summation; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.part1 = L"N"; obj.part2 = L"i=0"; obj.part3 = L"{}"; found = true;
                } else if (g_currentCommand == L"\\int") {
                    obj.type = MathType::Integral; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.part1 = L"b"; obj.part2 = L"a"; obj.part3 = L"{}"; found = true;
                } else if (g_currentCommand == L"\\sys") {
                    obj.type = MathType::SystemOfEquations; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.part1 = L""; obj.part2 = L""; obj.part3 = L""; found = true;
                } else if (g_currentCommand == L"\\sqrt") {
                    obj.type = MathType::SquareRoot; wcscpy_s(anchorStr, L"\u00A0\u00A0\u00A0\u00A0\u00A0");
                    obj.part1 = L""; obj.part2 = L""; obj.part3 = L""; found = true;
                }

                if (found) {
                    const LONG cmdLen = (LONG)g_currentCommand.size();
                    const LONG cmdStart = (LONG)selEndBefore - cmdLen;
                    if (cmdStart >= 0) {
                        wchar_t v[64] = {0}; TEXTRANGEW tr = { {cmdStart, (LONG)selEndBefore}, v };
                        SendMessage(hwnd, EM_GETTEXTRANGE, 0, (LPARAM)&tr);
                        if (g_currentCommand == v) {
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)cmdStart, (LPARAM)selEndBefore);
                            CHARFORMAT2W cf = {};
                            cf.cbSize = sizeof(cf);
                            SendMessage(hwnd, EM_GETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
                            LONG originalFontHeight = cf.yHeight;
                            // SquareRoot: normal-height anchors (radical drawn with GDI lines)
                            // Others: 2x-height anchors to reserve vertical space for limits
                            if (obj.type != MathType::SquareRoot) {
                                cf.dwMask |= CFM_SIZE; cf.yHeight *= 2;
                                SendMessage(hwnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);
                            }
                            SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)anchorStr);
                            // Push adjacent lines away from the math notation
                            {
                                PARAFORMAT2 pf2 = {};
                                pf2.cbSize = sizeof(pf2);
                                pf2.dwMask = PFM_SPACEBEFORE | PFM_SPACEAFTER;
                                if (obj.type == MathType::SquareRoot) {
                                    pf2.dySpaceBefore = (LONG)(originalFontHeight * 0.5);
                                    pf2.dySpaceAfter  = (LONG)(originalFontHeight * 0.3);
                                } else {
                                    pf2.dySpaceBefore = (LONG)(originalFontHeight * 1.5);
                                    pf2.dySpaceAfter  = (LONG)(originalFontHeight * 1.5);
                                }
                                SendMessage(hwnd, EM_SETPARAFORMAT, 0, (LPARAM)&pf2);
                            }
                            mgr.ShiftObjectsAfter(selEndBefore, 5 - cmdLen);
                            obj.barStart = cmdStart; obj.barLen = 5;
                            objects.push_back(std::move(obj));
                            state.objectIndex = objects.size() - 1; if (!state.active) HideCaret(hwnd);
                            state.active = true;
                            if (obj.type == MathType::SystemOfEquations) state.activePart = 1;
                            else if (obj.type == MathType::SquareRoot) state.activePart = 1;
                            else if (ch == L'^') state.activePart = 1; else if (ch == L'_') state.activePart = 2; else state.activePart = 3;
                            SendMessage(hwnd, EM_SETSEL, (WPARAM)(cmdStart + 5), (LPARAM)(cmdStart + 5));
                            InvalidateRect(hwnd, nullptr, TRUE); g_currentCommand.clear(); return 0;
                        }
                    }
                }
                g_currentCommand.clear();
            }

            if (ch == L'/' && !g_currentNumber.empty()) {
                const LONG nLen = (LONG)g_currentNumber.size();
                const LONG nS = (LONG)selEndBefore - nLen;
                const LONG bL = std::max<LONG>(3, nLen);
                SendMessage(hwnd, EM_SETSEL, (WPARAM)nS, (LPARAM)selEndBefore);
                SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)bL, L'\u2500').c_str());
                HideAnchorChars(hwnd, nS, bL);
                mgr.ShiftObjectsAfter(selEndBefore, bL - nLen);
                MathObject obj; obj.type = MathType::Fraction; obj.barStart = nS; obj.barLen = bL; obj.part1 = g_currentNumber;
                objects.push_back(std::move(obj));
                state.objectIndex = objects.size() - 1; if (!state.active) HideCaret(hwnd);
                state.active = true; state.activePart = 2;
                SendMessage(hwnd, EM_SETSEL, (WPARAM)(nS + bL), (LPARAM)(nS + bL));
                InvalidateRect(hwnd, nullptr, TRUE); g_currentNumber.clear(); return 0;
            }

            if (ch >= L'0' && ch <= L'9') { g_currentNumber += ch; g_currentCommand.clear(); }
            else if (ch == L'\\' || !g_currentCommand.empty()) { g_currentCommand += ch; g_currentNumber.clear(); }
            else { g_currentNumber.clear(); g_currentCommand.clear(); }

            LRESULT res = CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
            mgr.ShiftObjectsAfter((LONG)selEndBefore, (LONG)GetWindowTextLengthW(hwnd) - lenBefore);
            if (!objects.empty()) InvalidateRect(hwnd, nullptr, TRUE);
            return res;
        }
        }
        return CallWindowProc(g_originalProc, hwnd, uMsg, wParam, lParam);
    }
}

bool InstallMathSupport(HWND hRichEdit)
{
    if (!hRichEdit) return false;
    g_hEdit = hRichEdit; g_currentNumber.clear(); g_currentCommand.clear();
    MathManager::Get().Clear();
    if (!g_originalProc) g_originalProc = (WNDPROC)SetWindowLongPtr(g_hEdit, GWLP_WNDPROC, (LONG_PTR)MathRichEditProc);
    return !!g_originalProc;
}

void ResetMathSupport() { g_currentNumber.clear(); g_currentCommand.clear(); MathManager::Get().Clear(); if (g_hEdit) InvalidateRect(g_hEdit, nullptr, TRUE); }

void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator)
{
    if (!hEdit) return;
    DWORD s, e; SendMessage(hEdit, EM_GETSEL, (WPARAM)&s, (LPARAM)&e);
    const LONG bL = (LONG)std::max<size_t>(3, std::max(numerator.size(), denominator.size()));
    SendMessage(hEdit, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)std::wstring((size_t)bL, L'\u2500').c_str());
    HideAnchorChars(hEdit, (LONG)s, bL);
    MathObject obj; obj.type = MathType::Fraction; obj.barStart = (LONG)s; obj.barLen = bL; obj.part1 = numerator; obj.part2 = denominator;
    MathManager::Get().GetObjects().push_back(std::move(obj));
    MathManager::Get().ShiftObjectsAfter((LONG)s + 1, bL - (LONG)(e - s));
    SendMessage(hEdit, EM_SETSEL, (WPARAM)(s + bL), (LPARAM)(s + bL));
    InvalidateRect(hEdit, nullptr, TRUE);
}
