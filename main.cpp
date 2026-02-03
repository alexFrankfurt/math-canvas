#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <richedit.h>
#include <string>
#include <algorithm>

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK RichEditSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

// Global variables
HWND g_hRichEdit;
HWND g_hFractionButton;
HWND g_hClearButton;
std::wstring g_strCurrentNumber;

WNDPROC g_OriginalRichEditProc = nullptr;

struct FractionTypingState
{
    bool active = false;          // typing denominator (subscript)
    LONG numStart = 0;
    LONG numEnd = 0;
    LONG barStart = 0;
    LONG barLen = 0;
    LONG denomStart = 0;
    LONG minBarLen = 0;
    CHARFORMAT2W baseFormat = {}; // format to restore after finishing fraction
    bool hasBaseFormat = false;
};

FractionTypingState g_fractionState;

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

static CHARFORMAT2W MakeSuperscriptFormat(const CHARFORMAT2W& base)
{
    CHARFORMAT2W cf = base;
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_EFFECTS | CFM_OFFSET | CFM_SIZE;
    // Use baseline offset (yOffset) rather than CFE_SUPERSCRIPT; Rich Edit can ignore
    // custom yOffset when built-in super/sub effects are enabled.
    cf.dwEffects = base.dwEffects & ~(CFE_SUPERSCRIPT | CFE_SUBSCRIPT);
    // Bigger and raised; values are in twips (1/20 pt)
    cf.yHeight = std::max<LONG>(140, (base.yHeight * 8) / 10);
    cf.yOffset = std::max<LONG>(200, (base.yHeight * 8) / 10);
    return cf;
}

static CHARFORMAT2W MakeSubscriptFormat(const CHARFORMAT2W& base)
{
    CHARFORMAT2W cf = base;
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_EFFECTS | CFM_OFFSET | CFM_SIZE;
    cf.dwEffects = base.dwEffects & ~(CFE_SUPERSCRIPT | CFE_SUBSCRIPT);
    cf.yHeight = std::max<LONG>(140, (base.yHeight * 8) / 10);
    cf.yOffset = -std::max<LONG>(180, (base.yHeight * 6) / 10);
    return cf;
}

static CHARFORMAT2W MakeBarFormat(const CHARFORMAT2W& base)
{
    CHARFORMAT2W cf = MakeNormalFormat(base);
    cf.dwMask |= CFM_WEIGHT | CFM_SIZE;
    cf.wWeight = FW_BOLD;
    cf.yHeight = std::max<LONG>(base.yHeight, (base.yHeight * 11) / 10);
    return cf;
}

static void ApplyFormatToRange(HWND hEdit, LONG start, LONG end, const CHARFORMAT2W& cf)
{
    SendMessage(hEdit, EM_SETSEL, (WPARAM)start, (LPARAM)end);
    SendMessage(hEdit, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);
}

static void SetInsertionFormat(HWND hEdit, const CHARFORMAT2W& cf)
{
    // With an empty selection, SCF_SELECTION sets the insertion point format
    SendMessage(hEdit, EM_SETCHARFORMAT, (WPARAM)SCF_SELECTION, (LPARAM)&cf);
}

static void InsertFormattedFraction(HWND hEdit, const std::wstring& numerator, const std::wstring& denominator)
{
    DWORD selStart = 0, selEnd = 0;
    SendMessage(hEdit, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);

    CHARFORMAT2W base = GetSelectionCharFormat(hEdit);
    CHARFORMAT2W sup = MakeSuperscriptFormat(base);
    CHARFORMAT2W sub = MakeSubscriptFormat(base);
    CHARFORMAT2W normal = MakeNormalFormat(base);

    const LONG startPos = (LONG)selStart;

    SendMessage(hEdit, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)numerator.c_str());
    DWORD afterNumStart = 0, afterNumEnd = 0;
    SendMessage(hEdit, EM_GETSEL, (WPARAM)&afterNumStart, (LPARAM)&afterNumEnd);
    ApplyFormatToRange(hEdit, startPos, (LONG)afterNumStart, sup);

    // Move caret to end and insert a horizontal fraction bar (one or more characters)
    SendMessage(hEdit, EM_SETSEL, (WPARAM)afterNumStart, (LPARAM)afterNumStart);
    const size_t widest = (numerator.size() > denominator.size()) ? numerator.size() : denominator.size();
    const LONG barLen = (LONG)((widest < 3) ? 3 : widest);
    std::wstring bar(barLen, L'\u2500'); // BOX DRAWINGS LIGHT HORIZONTAL (looks like a fraction bar)
    CHARFORMAT2W barFmt = MakeBarFormat(base);
    SetInsertionFormat(hEdit, barFmt);
    SendMessage(hEdit, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)bar.c_str());

    DWORD afterSlashStart = 0, afterSlashEnd = 0;
    SendMessage(hEdit, EM_GETSEL, (WPARAM)&afterSlashStart, (LPARAM)&afterSlashEnd);
    SetInsertionFormat(hEdit, sub);
    SendMessage(hEdit, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)denominator.c_str());

    DWORD afterDenStart = 0, afterDenEnd = 0;
    SendMessage(hEdit, EM_GETSEL, (WPARAM)&afterDenStart, (LPARAM)&afterDenEnd);
    ApplyFormatToRange(hEdit, (LONG)afterSlashStart, (LONG)afterDenStart, sub);

    // Restore insertion format to normal
    SendMessage(hEdit, EM_SETSEL, (WPARAM)afterDenStart, (LPARAM)afterDenStart);
    SetInsertionFormat(hEdit, normal);
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
    // Load Rich Edit library
    HMODULE hRichEdit = LoadLibrary(L"Msftedit.dll");
    if (!hRichEdit)
    {
        MessageBox(nullptr, L"Failed to load Rich Edit library", L"Error", MB_OK | MB_ICONERROR);
        return 0;
    }
    
    // Register the window class
    const wchar_t CLASS_NAME[] = L"WinDeskAppWindowClass";

    WNDCLASS wc = {};
    wc.lpfnWndProc = WindowProc;
    wc.hInstance = hInstance;
    wc.lpszClassName = CLASS_NAME;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);

    RegisterClass(&wc);

    // Create the main window
    HWND hwnd = CreateWindowEx(
        0,
        CLASS_NAME,
        L"WinDeskApp - Text Editor with Fractions",
        WS_OVERLAPPEDWINDOW,
        CW_USEDEFAULT, CW_USEDEFAULT, 800, 600,
        nullptr,
        nullptr,
        hInstance,
        nullptr
    );

    if (hwnd == nullptr)
    {
        FreeLibrary(hRichEdit);
        return 0;
    }

    // Create Rich Edit control
    g_hRichEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        L"RICHEDIT50W",
        L"Type your text here...\n\nTwo-dimensional fraction feature:\n- Type a number, then '/' to create a fraction\n- Example: Type '3' then '/' becomes 3 over a line\n- Then type '4' to complete 3/4\n\nClick 'Clear Text' to reset.",
        WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | WS_VSCROLL | WS_HSCROLL | ES_WANTRETURN,
        10, 100, 760, 450,
        hwnd,
        nullptr,
        hInstance,
        nullptr
    );

    // Subclass Rich Edit so we can intercept '/' and format fractions
    g_OriginalRichEditProc = (WNDPROC)SetWindowLongPtr(g_hRichEdit, GWLP_WNDPROC, (LONG_PTR)RichEditSubclassProc);

    // Set font for Rich Edit control
    HFONT hFont = CreateFont(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    SendMessage(g_hRichEdit, WM_SETFONT, (WPARAM)hFont, TRUE);

    // Create fraction button
    g_hFractionButton = CreateWindow(
        L"BUTTON",
        L"Insert Fraction",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        10, 10, 150, 40,
        hwnd,
        (HMENU)1,
        hInstance,
        nullptr
    );

    // Create clear button
    g_hClearButton = CreateWindow(
        L"BUTTON",
        L"Clear Text",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        170, 10, 150, 40,
        hwnd,
        (HMENU)2,
        hInstance,
        nullptr
    );

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Run the message loop
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    FreeLibrary(hRichEdit);
    return 0;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_COMMAND:
        if (LOWORD(wParam) == 1) // Insert Fraction button
        {
            InsertFormattedFraction(g_hRichEdit, L"1", L"2");
        }
        else if (LOWORD(wParam) == 2) // Clear button
        {
            SetWindowText(g_hRichEdit, L"");
            g_strCurrentNumber.clear();
            g_fractionState = {};
        }
        return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        FillRect(hdc, &ps.rcPaint, (HBRUSH)(COLOR_WINDOW + 1));
        EndPaint(hwnd, &ps);
    }
    return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

LRESULT CALLBACK RichEditSubclassProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_KEYDOWN:
        // Navigation/edit keys reset digit collection; arrow keys also end fraction typing.
        if (wParam == VK_LEFT || wParam == VK_RIGHT || wParam == VK_UP || wParam == VK_DOWN ||
            wParam == VK_HOME || wParam == VK_END || wParam == VK_DELETE || wParam == VK_RETURN)
        {
            g_strCurrentNumber.clear();
            if (g_fractionState.active)
            {
                g_fractionState.active = false;
                if (g_fractionState.hasBaseFormat)
                {
                    CHARFORMAT2W normal = MakeNormalFormat(g_fractionState.baseFormat);
                    SetInsertionFormat(hwnd, normal);
                }
            }
        }
        break;

    case WM_CHAR:
    {
        const wchar_t ch = (wchar_t)wParam;

        // While typing denominator, keep subscript until a non-digit ends it.
        if (g_fractionState.active)
        {
            if (ch >= L'0' && ch <= L'9')
            {
                LRESULT res = CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);

                // If denominator grows longer than the bar, extend the bar and shift denom start.
                DWORD selStart = 0, selEnd = 0;
                SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);
                const LONG caretPos = (LONG)selEnd;
                const LONG denomLen = caretPos - g_fractionState.denomStart;
                if (denomLen > g_fractionState.barLen)
                {
                    const LONG extra = denomLen - g_fractionState.barLen;
                    const LONG insertPos = g_fractionState.barStart + g_fractionState.barLen;

                    CHARFORMAT2W barFmt = MakeBarFormat(g_fractionState.baseFormat);
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)insertPos, (LPARAM)insertPos);
                    SetInsertionFormat(hwnd, barFmt);
                    std::wstring barExtra((size_t)extra, L'\u2500');
                    SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)barExtra.c_str());

                    g_fractionState.barLen += extra;
                    g_fractionState.denomStart += extra;

                    // Restore caret (it shifted right by 'extra')
                    SendMessage(hwnd, EM_SETSEL, (WPARAM)(caretPos + extra), (LPARAM)(caretPos + extra));
                    CHARFORMAT2W sub = MakeSubscriptFormat(g_fractionState.baseFormat);
                    SetInsertionFormat(hwnd, sub);
                }

                return res;
            }

            if (ch == 0x08) // backspace
            {
                LRESULT res = CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);

                DWORD selStart = 0, selEnd = 0;
                SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);
                if ((LONG)selStart <= g_fractionState.denomStart)
                {
                    g_fractionState.active = false;
                    if (g_fractionState.hasBaseFormat)
                    {
                        CHARFORMAT2W normal = MakeNormalFormat(g_fractionState.baseFormat);
                        SetInsertionFormat(hwnd, normal);
                    }
                }

                return res;
            }

            // End denominator on any non-digit
            g_fractionState.active = false;
            if (g_fractionState.hasBaseFormat)
            {
                CHARFORMAT2W normal = MakeNormalFormat(g_fractionState.baseFormat);
                SetInsertionFormat(hwnd, normal);
            }
            g_strCurrentNumber.clear();
            return CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);
        }

        // Not in fraction mode
        if (ch >= L'0' && ch <= L'9')
        {
            g_strCurrentNumber += ch;
            return CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);
        }

        if (ch == 0x08) // backspace
        {
            if (!g_strCurrentNumber.empty())
            {
                g_strCurrentNumber.pop_back();
            }
            return CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);
        }

        if (ch == L'/')
        {
            if (!g_strCurrentNumber.empty())
            {
                DWORD selStart = 0, selEnd = 0;
                SendMessage(hwnd, EM_GETSEL, (WPARAM)&selStart, (LPARAM)&selEnd);
                const LONG caretPos = (LONG)selEnd;
                const LONG numLen = (LONG)g_strCurrentNumber.size();
                const LONG numStart = caretPos - numLen;
                const LONG numEnd = caretPos;

                CHARFORMAT2W base = GetSelectionCharFormat(hwnd);
                g_fractionState.baseFormat = base;
                g_fractionState.hasBaseFormat = true;

                CHARFORMAT2W sup = MakeSuperscriptFormat(base);
                CHARFORMAT2W sub = MakeSubscriptFormat(base);
                CHARFORMAT2W normal = MakeNormalFormat(base);
                CHARFORMAT2W barFmt = MakeBarFormat(base);

                // Format numerator (already inserted) as superscript
                ApplyFormatToRange(hwnd, numStart, numEnd, sup);

                // Move caret back to end of numerator and insert a horizontal bar
                SendMessage(hwnd, EM_SETSEL, (WPARAM)numEnd, (LPARAM)numEnd);
                const LONG barLen = std::max<LONG>(3, numLen);
                std::wstring bar((size_t)barLen, L'\u2500');
                SetInsertionFormat(hwnd, barFmt);
                SendMessage(hwnd, EM_REPLACESEL, (WPARAM)TRUE, (LPARAM)bar.c_str());

                DWORD afterBarStart = 0, afterBarEnd = 0;
                SendMessage(hwnd, EM_GETSEL, (WPARAM)&afterBarStart, (LPARAM)&afterBarEnd);

                g_fractionState.numStart = numStart;
                g_fractionState.numEnd = numEnd;
                g_fractionState.barStart = numEnd;
                g_fractionState.barLen = barLen;
                g_fractionState.minBarLen = barLen;
                g_fractionState.denomStart = numEnd + barLen;
                g_fractionState.active = true;
                SetInsertionFormat(hwnd, sub);

                g_strCurrentNumber.clear();
                return 0; // suppress default '/'
            }

            // No collected number; let '/' through
            g_strCurrentNumber.clear();
            return CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);
        }

        // Any other character resets digit collection
        g_strCurrentNumber.clear();
        return CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);
    }
    }

    return CallWindowProc(g_OriginalRichEditProc, hwnd, uMsg, wParam, lParam);
}
