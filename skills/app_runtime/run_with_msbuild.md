# Skill: Run Application using MSBuild and PowerShell

## Description
This skill explains how to build and launch a C++ Windows application (like `windeskapp`) directly from the command line using standard Microsoft developer tools. This is useful for automation, CI/CD, or quick testing without opening the Visual Studio IDE.

## Prerequisites
- **MSBuild.exe**: Usually located in a path like `C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe`.
- **PowerShell (pwsh)**: used for scripting the build and launch process.

## Steps

### 1) Build the project
Run `msbuild` against the project file (`.vcxproj`) or solution file (`.sln`). Specifying the configuration and platform ensures you get the expected output.

```powershell
# Build the Debug x64 version
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" .\CppProject.vcxproj /p:Configuration=Debug /p:Platform=x64 /m
```

### 2) Locate the output
By default, the output `.exe` is placed in a folder matching your configuration (e.g., `.\Debug\`).

```powershell
# Verify the file exists
Test-Path .\Debug\windeskapp.exe
```

### 3) Launch the application
Use `Start-Process` to run the executable. This launches it in a separate process, allowing your terminal to remain interactive.

```powershell
# Launch the app
Start-Process ".\Debug\windeskapp.exe"
```

### 4) One-liner (Build and Run)
You can chain these commands using a semicolon `;` to build and immediately run the app if the build succeeds.

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" .\CppProject.vcxproj /p:Configuration=Debug /p:Platform=x64 /m; Start-Process ".\Debug\windeskapp.exe"
```

## Troubleshooting
- **MSBuild not found**: Ensure the path to `MSBuild.exe` is correct for your Visual Studio installation (Community vs. Professional vs. Enterprise).
- **Build Errors**: Check the terminal output for compiler or linker errors.
- **App fails to start**: Refer to the "Start Application Without DLL Errors" skill for resolving runtime dependencies.

## Done Criteria
The application launches successfully after a successful build command, and you can interact with the window.
