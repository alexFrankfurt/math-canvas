#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <richedit.h>
#include <string>
#include <vector>

#include "src/math_editor.h"

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);

#ifdef _DEBUG
static bool RunFractionSelfTest(HWND hEdit, std::wstring& outDetails)
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
        std::vector<wchar_t> buf((size_t)originalLen + 1, L'\0');
        GetWindowTextW(hEdit, buf.data(), originalLen + 1);
        originalText.assign(buf.data());
    }

    // Start from a clean slate.
    SetWindowTextW(hEdit, L"");
    ResetMathSupport();
    SendMessage(hEdit, EM_SETSEL, 0, 0);

    // Simulate typing: 3/4
    SendMessage(hEdit, WM_CHAR, (WPARAM)L'3', 0);
    SendMessage(hEdit, WM_CHAR, (WPARAM)L'/', 0);
    SendMessage(hEdit, WM_CHAR, (WPARAM)L'4', 0);

    const int lenAfter = GetWindowTextLengthW(hEdit);
    std::wstring afterText;
    if (lenAfter > 0)
    {
        std::vector<wchar_t> buf((size_t)lenAfter + 1, L'\0');
        GetWindowTextW(hEdit, buf.data(), lenAfter + 1);
        afterText.assign(buf.data());
    }

    const bool hasBar = (afterText.find(L'\u2500') != std::wstring::npos);
    const bool hasSlash = (afterText.find(L'/') != std::wstring::npos);

    // Restore.
    SetWindowTextW(hEdit, originalText.c_str());
    ResetMathSupport();

    if (!hasBar)
    {
        outDetails = L"Expected U+2500 bar character to be inserted, but it was not.";
        return false;
    }
    if (hasSlash)
    {
        outDetails = L"Unexpected '/' remained in the RichEdit text (should be replaced by bar).";
        return false;
    }

    return true;
}
#endif

// Global variables
HWND g_hRichEdit;
HWND g_hFractionButton;
HWND g_hClearButton;
HWND g_hDarkModeButton;
bool g_darkMode = false;

static void ApplyTheme(HWND hwnd)
{
    COLORREF bkColor = g_darkMode ? RGB(30, 30, 30) : RGB(255, 255, 255);
    COLORREF textColor = g_darkMode ? RGB(220, 220, 220) : RGB(0, 0, 0);

    // Update Rich Edit
    SendMessage(g_hRichEdit, EM_SETBKGNDCOLOR, 0, (LPARAM)bkColor);

    CHARFORMAT2W cf = {};
    cf.cbSize = sizeof(cf);
    cf.dwMask = CFM_COLOR;
    cf.crTextColor = textColor;
    SendMessage(g_hRichEdit, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cf);
    SendMessage(g_hRichEdit, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM)&cf);

    // Force repaint
    InvalidateRect(hwnd, nullptr, TRUE);
    InvalidateRect(g_hRichEdit, nullptr, TRUE);
    
    // Update button text
    SetWindowTextW(g_hDarkModeButton, g_darkMode ? L"Light Mode" : L"Dark Mode");
}


static HMODULE LoadRichEditWithFallback(const wchar_t** outClassName)
{
    if (outClassName)
        *outClassName = nullptr;

    // Preferred: RichEdit 5.0 (available on modern Windows).
    if (HMODULE mod = LoadLibraryW(L"Msftedit.dll"))
    {
        if (outClassName)
            *outClassName = L"RICHEDIT50W";
        return mod;
    }

    // Fallback: RichEdit 2.0/3.0.
    if (HMODULE mod = LoadLibraryW(L"Riched20.dll"))
    {
        if (outClassName)
            *outClassName = L"RICHEDIT20W";
        return mod;
    }

    return nullptr;
}

int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, int nCmdShow)
{
    const wchar_t* richEditClass = nullptr;
    HMODULE hRichEdit = LoadRichEditWithFallback(&richEditClass);
    if (!hRichEdit || !richEditClass)
    {
        MessageBoxW(nullptr,
            L"Failed to load a Rich Edit library.\n\nTried:\n- Msftedit.dll (RICHEDIT50W)\n- Riched20.dll (RICHEDIT20W)",
            L"Error",
            MB_OK | MB_ICONERROR);
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
        richEditClass,
        L"Type your text here...\n\nTwo-dimensional fraction feature:\n- Type a number, then '/' to create a fraction\n- Example: Type '3' then '/' becomes 3 over a line\n- Then type '4' to complete 3/4\n\nClick 'Clear Text' to reset.",
        WS_VISIBLE | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | WS_VSCROLL | WS_HSCROLL | ES_WANTRETURN,
        10, 100, 760, 450,
        hwnd,
        nullptr,
        hInstance,
        nullptr
    );

    // Install fraction support (RichEdit subclass + overlay drawing)
    InstallMathSupport(g_hRichEdit);

    // Set font for Rich Edit control
    HFONT hFont = CreateFont(24, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Arial");
    SendMessage(g_hRichEdit, WM_SETFONT, (WPARAM)hFont, TRUE);

#ifdef _DEBUG
    // Quick correctness check: does typing 3/4 create a stacked fraction bar?
    // We reflect the result in the window title so it's visible without a debugger.
    {
        std::wstring details;
        const bool ok = RunFractionSelfTest(g_hRichEdit, details);
        if (ok)
        {
            SetWindowTextW(hwnd, L"WinDeskApp - Text Editor with Fractions (Fraction OK)");
        }
        else
        {
            SetWindowTextW(hwnd, L"WinDeskApp - Text Editor with Fractions (Fraction FAIL)");
            if (!details.empty())
                MessageBoxW(hwnd, details.c_str(), L"Fraction Self-Test Failed", MB_OK | MB_ICONERROR);
        }
    }
#endif

    // Create fraction button
    g_hFractionButton = CreateWindow(
        L"BUTTON",
        L"Insert &Fraction",
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
        L"&Clear Text",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        170, 10, 150, 40,
        hwnd,
        (HMENU)2,
        hInstance,
        nullptr
    );

    // Create dark mode button
    g_hDarkModeButton = CreateWindow(
        L"BUTTON",
        L"Dark Mode",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        330, 10, 150, 40,
        hwnd,
        (HMENU)3,
        hInstance,
        nullptr
    );

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Initial theme
    ApplyTheme(hwnd);

    // Make sure typing goes to the editor immediately.
    SetFocus(g_hRichEdit);


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
            SetFocus(g_hRichEdit);
        }
        else if (LOWORD(wParam) == 2) // Clear button
        {
            SetWindowText(g_hRichEdit, L"");
            ResetMathSupport();
            SetFocus(g_hRichEdit);
        }
        else if (LOWORD(wParam) == 3) // Dark Mode button
        {
            g_darkMode = !g_darkMode;
            ApplyTheme(hwnd);
        }
        return 0;

    case WM_SIZE:

    {
        const int width = LOWORD(lParam);
        const int height = HIWORD(lParam);

        const int margin = 10;
        const int top = 100; // Leave space for buttons and instructions

        if (g_hRichEdit)
        {
            int editWidth = width - 2 * margin;
            int editHeight = height - top - margin;

            if (editWidth < 0) editWidth = 0;
            if (editHeight < 0) editHeight = 0;

            MoveWindow(
                g_hRichEdit,
                margin,
                top,
                editWidth,
                editHeight,
                TRUE);
        }
    }
    return 0;

    case WM_DESTROY:
        PostQuitMessage(0);
        return 0;

    case WM_PAINT:
    {
        PAINTSTRUCT ps;
        HDC hdc = BeginPaint(hwnd, &ps);
        HBRUSH hbr = CreateSolidBrush(g_darkMode ? RGB(45, 45, 45) : GetSysColor(COLOR_BTNFACE));
        FillRect(hdc, &ps.rcPaint, hbr);
        DeleteObject(hbr);
        EndPaint(hwnd, &ps);
    }
    return 0;
    }

    return DefWindowProc(hwnd, uMsg, wParam, lParam);
}
