# WinDeskApp

WinDeskApp is a Windows desktop math editor built on top of a RichEdit control with GDI overlay rendering for structured notation. It lets you mix normal text with editable math objects so fractions, roots, powers, logs, matrices, determinants, systems of equations, and nested expressions stay semantic instead of collapsing into plain text.

## Showcase
[watch](https://www.youtube.com/watch?v=NCobAOosHg8)

## What the app currently supports

- Fractions from either `number/` or `\frac` + `Space`
- Square roots with optional index via `\sqrt` + `Space`
- Absolute value, powers, and logarithms via `\abs`, `\pow`, and `\log`
- Summation, product, and integral templates with editable limits via `\sum`, `\prod`, and `\int`
- 2x2 matrices and determinants via `\mat` and `\det`
- Systems of equations via `\sys`
- Expression objects and function templates via `\expr`, `\sin`, `\cos`, `\tan`, `\asin`, `\acos`, `\atan`, `\ln`, and `\exp`
- True nested math inside active slots for square roots, fractions, powers, absolute values, and logarithms
- Structured copy, cut, paste, and `.wdm` document persistence so nested objects survive round trips
- Unit-aware evaluation with an inline unit suggestion popup while editing math

## Architecture at a glance

- `main.cpp`: Win32 shell, toolbar/menu wiring, document open/save flow, dirty-state prompts, and app startup checks
- `src/math_editor.cpp`: RichEdit subclassing, command expansion, math editing, nested caret/navigation logic, clipboard, and unit suggestion popup
- `src/math_renderer.cpp`: measurement, drawing, overlay caret geometry, and hit-testing for structured math
- `src/math_manager.cpp`: math-object management and formatted result generation
- `src/math_evaluator.cpp`: expression evaluation, system solving, determinant evaluation, and unit-aware arithmetic
- `src/math_types.h`: structured math model, slot/node helpers, and semantic serialization helpers

The core design is RichEdit text plus anchor-backed math objects. RichEdit owns the text flow, while the math renderer draws and hit-tests structured notation over the anchored positions.

## Requirements

- Windows
- Visual Studio 2022 with the Desktop development with C++ workload
- MSVC toolset `v143`
- RichEdit available through `Msftedit.dll`
- AutoHotkey v2 if you want to run the GUI smoke scripts in `ahk_tools/`

The checked-in solution and project files currently target `Debug|x64`.

## Build and run

The main application project is `CppProject.vcxproj`. The solution `windeskapp.sln` currently includes that app project.

### Build with MSBuild

```powershell
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
& $msbuild .\windeskapp.sln /p:Configuration=Debug /p:Platform=x64 /m
```

Primary output:

- `Debug\windeskapp.exe`

The app project also copies a convenience `windeskapp.exe` into the repo root after a successful build.

### Launch the app

```powershell
Start-Process .\Debug\windeskapp.exe
```

Or use the helper script:

```powershell
.\run_windeskapp.ps1
```

If you hit a missing debug runtime DLL such as `VCRUNTIME140D.dll` or `ucrtbased.dll`, you are trying to run a debug build without the MSVC debug runtime available. In that case, run from a development machine with Visual Studio installed, or add a Release configuration before distributing binaries.

## Typing commands in the editor

Type a command and press `Space` to expand it.

| Input | Result |
| --- | --- |
| `number/` or `\frac` | Fraction |
| `\sqrt` | Square root |
| `\abs` | Absolute value |
| `\pow` | Power / exponent |
| `\log` | Logarithm |
| `\sum` | Summation with limits |
| `\prod` | Product with limits |
| `\int` | Integral with limits |
| `\mat` | 2x2 matrix |
| `\det` | 2x2 determinant |
| `\sys` | System of equations |
| `\expr` | Expression object |
| `\sin`, `\cos`, `\tan`, `\asin`, `\acos`, `\atan`, `\ln`, `\exp` | Function template |

Useful editing behavior:

- Press `=` to evaluate supported expressions and determinants
- While editing a math object, type `\sqrt`, `\frac`, `\pow`, `\abs`, or `\log` inside an active slot and press `Space` to nest another structured object
- Press `Right` to enter the first nested slot
- Press `Left` or `Home` to return to the parent slot
- Press `Tab` to move across sibling slots such as fraction numerator/denominator or matrix cells
- In a square root, press `_` or `Tab` to move into the optional index slot
- Use `Ctrl+O` and `Ctrl+S` or the `File` menu for document operations

## Structured documents and clipboard

WinDeskApp stores structured documents as UTF-8 `.wdm` files. The document snapshot format preserves:

- plain RichEdit text
- anchor positions for math objects
- semantic payloads for structured math, including nested content

The same semantic model is also used for clipboard round trips, so copy/paste can preserve structured objects instead of falling back to plain text only.

## Testing and automation

Focused regression targets exist for the structured math model and document persistence.

### Build and run the core regression targets

```powershell
$msbuild = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

& $msbuild .\test_math_model.vcxproj /p:Configuration=Debug /p:Platform=x64 /m
& .\Debug\test_math_model.exe

& $msbuild .\test_document_persistence.vcxproj /p:Configuration=Debug /p:Platform=x64 /m
& .\Debug\test_document_persistence.exe
```

Other test and verification entry points in the repo:

- `test_eval.cpp`: evaluator-focused checks
- `test_linear_system.cpp` plus `build_and_test.bat`: lightweight linear-system test path
- `test_system_equation.cpp` and `test_system_equation_expanded.cpp`: system-equation experiments and validation helpers
- `ahk_tools/`: AutoHotkey v2 smoke scripts for live UI verification, including nested math, alignment, screenshot capture, equality evaluation, and unit dropdown behavior

Useful AHK scripts include:

- `ahk_tools/demo_showcase.ahk`
- `ahk_tools/test_nested_math.ahk`
- `ahk_tools/test_unit_dropdown.ahk`
- `ahk_tools/test_align.ahk`
- `ahk_tools/take_screenshot.ahk`

## Project status and known gaps

The repository already contains working structured nested math, structured clipboard/document persistence, dirty-state aware open/save flows, and a growing automated regression surface. The main remaining work is polish and cleanup rather than first-time feature enablement.

Known gaps called out in the repo docs today:

- Mixed baseline alignment and spacing for some nested layouts still need more manual visual verification
- Deeper repaint/clipping/flicker scenarios still need more live verification
- Matrix support is currently centered on structured 2x2 editing and determinant evaluation, not general matrix algebra
- Compatibility cleanup around legacy mirrored part fields is still in progress

For the current implementation roadmap and manual verification notes, see:

- `NESTED_MATH_IMPLEMENTATION_CHECKLIST.md`
- `NESTED_MATH_VERIFICATION_NOTES.md`

## Repository layout

```text
.
|- main.cpp
|- src/
|  |- math_editor.cpp
|  |- math_renderer.cpp
|  |- math_manager.cpp
|  |- math_evaluator.cpp
|  |- math_types.h
|- ahk_tools/
|- test_math_model.cpp
|- test_document_persistence.cpp
|- test_eval.cpp
|- test_linear_system.cpp
|- NESTED_MATH_IMPLEMENTATION_CHECKLIST.md
`- NESTED_MATH_VERIFICATION_NOTES.md
```

## Notes for contributors

- Build changes against `Debug|x64` unless you also update the checked-in project configuration set
- New notation features usually require coordinated changes in `src/math_editor.cpp`, `src/math_renderer.cpp`, and `src/math_manager.cpp`
- If you change the structured math model, also check the regression surface in `test_math_model.cpp` and `test_document_persistence.cpp`
- If you change nested layout or interaction behavior, update the verification notes and relevant AHK smoke scripts
