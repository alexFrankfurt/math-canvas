# Skill: Start Application Without DLL Errors (Win32 / MSVC)

## Description
This skill helps you launch a Windows desktop app (like `windeskapp.exe`) and resolve common startup failures such as:
- “The code execution cannot proceed because <dll> was not found”
- “0xc000007b” (architecture / bad image)
- “MSVCP140*.dll / VCRUNTIME140*.dll missing” (VC++ runtime)
- “VCRUNTIME140D.dll / ucrtbased.dll missing” (Debug runtime)

## Quick Triage (fastest path)
1. **Read the exact DLL name** from the dialog (copy it).
2. **Decide Debug vs Release**:
   - If the missing DLL ends with **`D.dll`** (example: `VCRUNTIME140D.dll`, `MSVCP140D.dll`, `ucrtbased.dll`) → you’re trying to run a **Debug** build on a machine without the Debug CRT.
   - If it ends without `D` (example: `VCRUNTIME140.dll`, `MSVCP140.dll`) → it’s usually the **VC++ Redistributable**.
3. **Check x64 vs x86** (0xc000007b is often this): confirm the app and the runtime you installed match architecture.

## Steps
### 1) Confirm what you’re running
- Prefer launching from PowerShell so the path is explicit:
  - `Start-Process .\Debug\windeskapp.exe`
- Verify the file exists:
  - `Test-Path .\Debug\windeskapp.exe`

### 2) If it’s a Debug runtime DLL (`*D.dll`)
**Symptom**: Missing `VCRUNTIME140D.dll`, `MSVCP140D.dll`, `ucrtbased.dll`.

**Fix options** (pick one):
1. **Build and run Release** instead of Debug.
   - In Visual Studio: Configuration = Release, Platform = x64.
2. **Run from Visual Studio** on the dev machine (it supplies the Debug CRT).
3. **Install Visual Studio / C++ build tools** on the target machine (dev-only scenario).

Notes:
- Debug CRT DLLs are not meant to be shipped to end users.

### 3) If it’s a Release VC++ runtime DLL (`VCRUNTIME140*.dll`, `MSVCP140*.dll`)
**Fix**:
- Install the **Microsoft Visual C++ Redistributable (VS 2015–2022)** matching your app’s architecture.
  - For this repo’s default build (`/p:Platform=x64`), install the **x64** redist.

### 4) If error is `0xc000007b` (Bad Image)
This commonly means **bitness mismatch** (x64 app trying to load x86 DLL or vice versa).

**Fix checklist**:
1. Confirm you built **x64** and are running the **x64** executable.
2. Ensure the installed VC++ redist matches that architecture.
3. If you have any app-local DLLs (next to the `.exe`), confirm they are the same architecture.

### 5) If the missing DLL is a Windows component (e.g., `msftedit.dll`)
This repo loads RichEdit via:
- `LoadLibrary(L"Msftedit.dll")`

**Fix options**:
1. **Prefer running on a modern Windows version** where `Msftedit.dll` exists in system folders.
2. If you must support older systems, implement a fallback:
   - Try `Msftedit.dll` and class `RICHEDIT50W`
   - If that fails, try `Riched20.dll` and class `RICHEDIT20W`

(That fallback is a code change, but it eliminates “missing Msftedit.dll” startup failures.)

### 6) Identify dependencies when the dialog isn’t clear
Use one of these approaches:
- **Dependencies** (GUI tool) to open the EXE and see missing imports.
- Visual Studio tools (if installed):
  - `dumpbin /dependents path\to\windeskapp.exe`

## Common Issues and Solutions
- **Running `Debug\windeskapp.exe` on another PC fails**: build **Release** or install the VC++ redist.
- **Works on dev machine only**: usually missing VC++ runtime on the target machine.
- **0xc000007b after installing redist**: installed x86 redist but running x64 app (or vice versa).
- **RichEdit DLL load error**: `Msftedit.dll` missing → use a fallback (`Riched20.dll`) or modern Windows.

## Repo-specific Notes (windeskapp)
- The build is currently configured for `Debug|x64` in the solution/project.
- Output is typically in `Debug\windeskapp.exe`.
- The app already shows a message box if `Msftedit.dll` fails to load.

## Done Criteria
You can reliably launch the app via `Start-Process` with no “missing DLL” dialogs, and (if applicable) the correct RichEdit library loads.
