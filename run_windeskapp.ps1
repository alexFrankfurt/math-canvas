# PowerShell script to build and run windeskapp
param(
    [string]$Configuration = "Debug",
    [string]$Platform = "x64"
)

$MSBuildPath = "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"

Write-Host "Building windeskapp ($Configuration|$Platform)..." -ForegroundColor Yellow

# Build the project
& $MSBuildPath .\CppProject.vcxproj /p:Configuration=$Configuration /p:Platform=$Platform /m /nologo

if ($LASTEXITCODE -eq 0) {
    Write-Host "Build successful!" -ForegroundColor Green
    
    $exePath = ".\$Configuration\windeskapp.exe"
    if (Test-Path $exePath) {
        Write-Host "Launching windeskapp..." -ForegroundColor Cyan
        Start-Process $exePath
        Write-Host "Application started successfully!" -ForegroundColor Green
    } else {
        Write-Host "Error: Executable not found at $exePath" -ForegroundColor Red
    }
} else {
    Write-Host "Build failed with exit code $LASTEXITCODE" -ForegroundColor Red
}