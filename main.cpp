#ifndef NOMINMAX
#define NOMINMAX
#endif
#include <windows.h>
#include <commdlg.h>
#include <richedit.h>
#include <string>
#include <vector>

#include "src/math_editor.h"
#include "src/math_manager.h"

// Forward declarations
LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
static void ApplyTheme(HWND hwnd);
static void UpdateWindowTitle(HWND hwnd);

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
HWND g_hOpenButton;
HWND g_hSaveButton;
bool g_darkMode = false;
HACCEL g_hAccelerators = nullptr;
std::wstring g_currentDocumentPath;
std::wstring g_startupStatusSuffix;
bool g_isDocumentDirty = false;
bool g_suppressDirtyTracking = false;

namespace
{
    constexpr const wchar_t* kAppTitleBase = L"WinDeskApp";
    constexpr int kOpenButtonId = 1;
    constexpr int kClearButtonId = 2;
    constexpr int kDarkModeButtonId = 3;
    constexpr int kFractionButtonId = 4;
    constexpr int kSaveButtonId = 5;
    constexpr int kSaveAsCommandId = 6;

    static std::wstring GetDisplayDocumentName()
    {
        if (g_currentDocumentPath.empty())
            return L"Untitled";

        const size_t slash = g_currentDocumentPath.find_last_of(L"\\/");
        if (slash == std::wstring::npos)
            return g_currentDocumentPath;
        return g_currentDocumentPath.substr(slash + 1);
    }

    static void SetDocumentDirty(HWND hwnd, bool dirty)
    {
        g_isDocumentDirty = dirty;
        if (g_hRichEdit)
            SendMessage(g_hRichEdit, EM_SETMODIFY, dirty ? TRUE : FALSE, 0);
        UpdateWindowTitle(hwnd);
    }

    static void SetCurrentDocumentPath(HWND hwnd, const std::wstring& path)
    {
        g_currentDocumentPath = path;
        UpdateWindowTitle(hwnd);
    }

    static HMENU BuildAppMenuBar()
    {
        HMENU menuBar = CreateMenu();
        HMENU fileMenu = CreatePopupMenu();
        if (!menuBar || !fileMenu)
            return nullptr;

        AppendMenuW(fileMenu, MF_STRING, (UINT_PTR)kOpenButtonId, L"&Open...\tCtrl+O");
        AppendMenuW(fileMenu, MF_STRING, (UINT_PTR)kSaveButtonId, L"&Save\tCtrl+S");
        AppendMenuW(fileMenu, MF_STRING, (UINT_PTR)kSaveAsCommandId, L"Save &As...");
        AppendMenuW(menuBar, MF_POPUP, (UINT_PTR)fileMenu, L"&File");
        return menuBar;
    }

    static HACCEL BuildAccelerators()
    {
        ACCEL accelerators[] = {
            { FCONTROL | FVIRTKEY, (WORD)'O', (WORD)kOpenButtonId },
            { FCONTROL | FVIRTKEY, (WORD)'S', (WORD)kSaveButtonId },
        };
        return CreateAcceleratorTableW(accelerators, (int)(sizeof(accelerators) / sizeof(accelerators[0])));
    }

    static bool WriteUtf8File(const std::wstring& path, const std::wstring& text)
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

    static bool ReadUtf8File(const std::wstring& path, std::wstring& outText)
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

    static bool PromptForDocumentPath(HWND owner, bool save, std::wstring& outPath)
    {
        wchar_t filePath[MAX_PATH] = L"";
        OPENFILENAMEW ofn = {};
        ofn.lStructSize = sizeof(ofn);
        ofn.hwndOwner = owner;
        ofn.lpstrFilter = L"WinDeskApp Math Documents (*.wdm)\0*.wdm\0All Files (*.*)\0*.*\0\0";
        ofn.lpstrFile = filePath;
        ofn.nMaxFile = MAX_PATH;
        ofn.lpstrDefExt = L"wdm";
        ofn.Flags = OFN_EXPLORER | OFN_PATHMUSTEXIST | OFN_HIDEREADONLY;
        if (save)
            ofn.Flags |= OFN_OVERWRITEPROMPT;
        else
            ofn.Flags |= OFN_FILEMUSTEXIST;

        const BOOL result = save ? GetSaveFileNameW(&ofn) : GetOpenFileNameW(&ofn);
        if (!result)
            return false;

        outPath = filePath;
        return true;
    }

    static bool SaveMathDocumentToPath(HWND owner, HWND hEdit, const std::wstring& path, std::wstring& outError)
    {
        outError.clear();
        std::wstring payload;
        if (!SerializeMathDocument(hEdit, payload))
        {
            outError = L"Failed to serialize the current document.";
            return false;
        }
        if (!WriteUtf8File(path, payload))
        {
            outError = L"Failed to write the document file.";
            return false;
        }
        return true;
    }

    static bool LoadMathDocumentFromPath(HWND owner, HWND hEdit, const std::wstring& path, std::wstring& outError)
    {
        outError.clear();
        std::wstring payload;
        if (!ReadUtf8File(path, payload))
        {
            outError = L"Failed to read the document file.";
            return false;
        }
        if (!TryDeserializeMathDocument(hEdit, payload))
        {
            outError = L"The file did not contain a valid structured math document.";
            return false;
        }
        return true;
    }

    static bool SaveDocumentInteractive(HWND hwnd, bool saveAs, bool& outCanceled)
    {
        outCanceled = false;

        std::wstring path = g_currentDocumentPath;
        if (saveAs || path.empty())
        {
            if (!PromptForDocumentPath(hwnd, true, path))
            {
                outCanceled = true;
                return false;
            }
        }

        std::wstring error;
        if (!SaveMathDocumentToPath(hwnd, g_hRichEdit, path, error))
        {
            MessageBoxW(hwnd, error.c_str(), L"Save Document Failed", MB_OK | MB_ICONERROR);
            return false;
        }

        SetCurrentDocumentPath(hwnd, path);
        SetDocumentDirty(hwnd, false);
        SetFocus(g_hRichEdit);
        return true;
    }

    static bool ConfirmDiscardUnsavedChanges(HWND hwnd)
    {
        if (!g_isDocumentDirty)
            return true;

        const std::wstring prompt = L"Save changes to " + GetDisplayDocumentName() + L" before continuing?";
        const int result = MessageBoxW(hwnd, prompt.c_str(), L"Unsaved Changes", MB_YESNOCANCEL | MB_ICONWARNING);
        if (result == IDCANCEL)
            return false;
        if (result == IDNO)
            return true;

        bool canceled = false;
        return SaveDocumentInteractive(hwnd, false, canceled);
    }

    static bool HandleAppCommand(HWND hwnd, int commandId)
    {
        if (commandId == kOpenButtonId)
        {
            if (!ConfirmDiscardUnsavedChanges(hwnd))
                return true;

            std::wstring path;
            if (PromptForDocumentPath(hwnd, false, path))
            {
                std::wstring error;
                g_suppressDirtyTracking = true;
                const bool loaded = LoadMathDocumentFromPath(hwnd, g_hRichEdit, path, error);
                g_suppressDirtyTracking = false;
                if (!loaded)
                    MessageBoxW(hwnd, error.c_str(), L"Open Document Failed", MB_OK | MB_ICONERROR);
                else
                {
                    SetCurrentDocumentPath(hwnd, path);
                    SetDocumentDirty(hwnd, false);
                    SetFocus(g_hRichEdit);
                }
            }
            return true;
        }

        if (commandId == kSaveButtonId)
        {
            bool canceled = false;
            SaveDocumentInteractive(hwnd, false, canceled);
            return true;
        }

        if (commandId == kSaveAsCommandId)
        {
            bool canceled = false;
            SaveDocumentInteractive(hwnd, true, canceled);
            return true;
        }

        if (commandId == kFractionButtonId)
        {
            InsertFormattedFraction(g_hRichEdit, L"1", L"2");
            SetFocus(g_hRichEdit);
            return true;
        }

        if (commandId == kClearButtonId)
        {
            SetWindowText(g_hRichEdit, L"");
            ResetMathSupport();
            SetFocus(g_hRichEdit);
            return true;
        }

        if (commandId == kDarkModeButtonId)
        {
            g_darkMode = !g_darkMode;
            ApplyTheme(hwnd);
            return true;
        }

        return false;
    }

#ifdef _DEBUG
    static bool DebugRunStructuredDocumentFileRoundTripSelfTest(HWND hEdit, std::wstring& outDetails)
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

        SetWindowTextW(hEdit, L"");
        ResetMathSupport();
        SendMessage(hEdit, EM_SETSEL, 0, 0);
        SendMessage(hEdit, WM_CHAR, (WPARAM)L'2', 0);
        SendMessage(hEdit, WM_CHAR, (WPARAM)L'/', 0);
        SendMessage(hEdit, WM_CHAR, (WPARAM)L'5', 0);
        SendMessage(hEdit, WM_KEYDOWN, VK_RETURN, 0);
        SendMessage(hEdit, WM_CHAR, (WPARAM)L'X', 0);

        wchar_t tempPath[MAX_PATH] = L"";
        wchar_t tempFile[MAX_PATH] = L"";
        if (!GetTempPathW(MAX_PATH, tempPath) || !GetTempFileNameW(tempPath, L"wdm", 0, tempFile))
        {
            outDetails = L"Failed to allocate a temp file for file round-trip self-test.";
            SetWindowTextW(hEdit, originalText.c_str());
            ResetMathSupport();
            return false;
        }

        std::wstring error;
        const bool saveOk = SaveMathDocumentToPath(nullptr, hEdit, tempFile, error);
        if (saveOk)
        {
            SetWindowTextW(hEdit, L"");
            ResetMathSupport();
        }
        const bool loadOk = saveOk && LoadMathDocumentFromPath(nullptr, hEdit, tempFile, error);
        DeleteFileW(tempFile);

        if (!saveOk || !loadOk)
        {
            outDetails = error.empty() ? L"File round-trip self-test failed." : error;
            SetWindowTextW(hEdit, originalText.c_str());
            ResetMathSupport();
            return false;
        }

        std::wstring payloadAfterLoad;
        if (!SerializeMathDocument(hEdit, payloadAfterLoad))
        {
            outDetails = L"Failed to reserialize the file-loaded document during self-test.";
            SetWindowTextW(hEdit, originalText.c_str());
            ResetMathSupport();
            return false;
        }

        if (MathManager::Get().GetObjects().size() != 1)
        {
            outDetails = L"Expected one structured object after file round-trip self-test.";
            SetWindowTextW(hEdit, originalText.c_str());
            ResetMathSupport();
            return false;
        }

        SetWindowTextW(hEdit, originalText.c_str());
        ResetMathSupport();
        return true;
    }
#endif
}

static void UpdateWindowTitle(HWND hwnd)
{
    if (!hwnd)
        return;

    std::wstring title = std::wstring(kAppTitleBase) + L" - " + GetDisplayDocumentName();
    if (g_isDocumentDirty)
        title += L"*";
    if (!g_startupStatusSuffix.empty())
        title += g_startupStatusSuffix;
    SetWindowTextW(hwnd, title.c_str());
}

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

    HMENU hMenuBar = BuildAppMenuBar();
    if (hMenuBar)
        SetMenu(hwnd, hMenuBar);
    g_hAccelerators = BuildAccelerators();

    // Create Rich Edit control
    g_hRichEdit = CreateWindowEx(
        WS_EX_CLIENTEDGE,
        richEditClass,
        L"Type your text here...\n\nMath features:\n- Type a number, then '/' to create a fraction, or type \\frac then Space\n- Type \\sqrt then Space for a square root; press '_' or Tab to enter an optional index\n- Type \\abs then Space to insert absolute value\n- Type \\pow then Space to insert power (base^exponent)\n- Type \\log then Space to insert logarithm\n- Type \\sum, \\prod, or \\int then Space to insert operators with limits\n- Type \\mat then Space to insert a 2x2 matrix; Tab moves across the four cells\n- Type \\det then Space to insert a 2x2 determinant and evaluate it with '='\n- Type \\sys then Space to insert a system of equations\n- Type \\expr then Space to enter an expression and evaluate it with '='\n- Type \\sin, \\cos, \\tan, \\asin, \\acos, \\atan, \\ln, or \\exp then Space for a function template\n- While editing math, type \\sqrt, \\frac, \\pow, \\abs, or \\log inside an active part and press Space to nest it\n- Press Right to enter the first nested slot, Left or Home to return to the parent slot, Tab to move across slots\n- Use File > Open / Save / Save As or Ctrl+O / Ctrl+S to load and store structured .wdm documents\n- Opening another file or closing the window will prompt to save unsaved structured edits first\n\nClick 'Clear Text' to reset.",
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
    ApplyMathEditInsets(g_hRichEdit);

#ifdef _DEBUG
    // Quick correctness check: does typing 3/4 create a stacked fraction bar?
    // We reflect the result in the window title so it's visible without a debugger.
    {
        g_suppressDirtyTracking = true;
        std::wstring fractionDetails;
        const bool fractionOk = RunFractionSelfTest(g_hRichEdit, fractionDetails);
        std::wstring roundTripDetails;
        const bool roundTripOk = DebugRunStructuredRoundTripSelfTest(g_hRichEdit, roundTripDetails);
        std::wstring fragmentDetails;
        const bool fragmentOk = DebugRunStructuredFragmentRoundTripSelfTest(g_hRichEdit, fragmentDetails);
        std::wstring documentDetails;
        const bool documentOk = DebugRunStructuredDocumentRoundTripSelfTest(g_hRichEdit, documentDetails);
        std::wstring documentFileDetails;
        const bool documentFileOk = DebugRunStructuredDocumentFileRoundTripSelfTest(g_hRichEdit, documentFileDetails);
        if (fractionOk && roundTripOk && fragmentOk && documentOk && documentFileOk)
        {
            g_startupStatusSuffix = L" (Fraction OK, RoundTrip OK, Fragment OK, Document OK, File OK)";
        }
        else
        {
            g_startupStatusSuffix = L" (Self-Test FAIL)";
            std::wstring message;
            if (!fractionOk)
                message += L"Fraction self-test failed: " + fractionDetails + L"\n\n";
            if (!roundTripOk)
                message += L"Structured round-trip self-test failed: " + roundTripDetails;
            if (!fragmentOk)
            {
                if (!message.empty()) message += L"\n\n";
                message += L"Structured fragment self-test failed: " + fragmentDetails;
            }
            if (!documentOk)
            {
                if (!message.empty()) message += L"\n\n";
                message += L"Structured document self-test failed: " + documentDetails;
            }
            if (!documentFileOk)
            {
                if (!message.empty()) message += L"\n\n";
                message += L"Structured document file self-test failed: " + documentFileDetails;
            }
            if (!message.empty())
                MessageBoxW(hwnd, message.c_str(), L"Math Self-Test Failed", MB_OK | MB_ICONERROR);
        }
        g_suppressDirtyTracking = false;
    }
#else
    g_startupStatusSuffix.clear();
#endif

    g_hOpenButton = CreateWindow(
        L"BUTTON",
        L"&Open",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        10, 10, 100, 40,
        hwnd,
        (HMENU)(INT_PTR)kOpenButtonId,
        hInstance,
        nullptr
    );

    g_hSaveButton = CreateWindow(
        L"BUTTON",
        L"&Save",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        120, 10, 100, 40,
        hwnd,
        (HMENU)(INT_PTR)kSaveButtonId,
        hInstance,
        nullptr
    );

    // Create fraction button
    g_hFractionButton = CreateWindow(
        L"BUTTON",
        L"Insert &Fraction",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_DEFPUSHBUTTON,
        230, 10, 150, 40,
        hwnd,
        (HMENU)(INT_PTR)kFractionButtonId,
        hInstance,
        nullptr
    );

    // Create clear button
    g_hClearButton = CreateWindow(
        L"BUTTON",
        L"&Clear Text",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        390, 10, 150, 40,
        hwnd,
        (HMENU)(INT_PTR)kClearButtonId,
        hInstance,
        nullptr
    );

    // Create dark mode button
    g_hDarkModeButton = CreateWindow(
        L"BUTTON",
        L"Dark Mode",
        WS_TABSTOP | WS_VISIBLE | WS_CHILD | BS_PUSHBUTTON,
        550, 10, 150, 40,
        hwnd,
        (HMENU)(INT_PTR)kDarkModeButtonId,
        hInstance,
        nullptr
    );

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    // Initial theme
    ApplyTheme(hwnd);
    SetCurrentDocumentPath(hwnd, L"");
    SetDocumentDirty(hwnd, false);

    // Make sure typing goes to the editor immediately.
    SetFocus(g_hRichEdit);


    // Run the message loop
    MSG msg = {};
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (g_hAccelerators && TranslateAcceleratorW(hwnd, g_hAccelerators, &msg))
            continue;
        TranslateMessage(&msg);
        DispatchMessage(&msg);
    }

    if (g_hAccelerators)
        DestroyAcceleratorTable(g_hAccelerators);
    FreeLibrary(hRichEdit);
    return 0;
}

LRESULT CALLBACK WindowProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    switch (uMsg)
    {
    case WM_COMMAND:
        if (HIWORD(wParam) == EN_CHANGE && (HWND)lParam == g_hRichEdit)
        {
            if (!g_suppressDirtyTracking)
                SetDocumentDirty(hwnd, true);
            return 0;
        }
        if (HandleAppCommand(hwnd, LOWORD(wParam)))
            return 0;
        break;

    case WM_CLOSE:
        if (!ConfirmDiscardUnsavedChanges(hwnd))
            return 0;
        DestroyWindow(hwnd);
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
            ApplyMathEditInsets(g_hRichEdit);
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
