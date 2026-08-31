# ============================================
# EstimateTool Build Script
# ============================================

Write-Host "============================================"
Write-Host "EstimateTool Build Start"
Write-Host "============================================"


# --------------------------------------------
# 1. Clean
# --------------------------------------------

Write-Host ""
Write-Host "[1/5] Cleaning old build files..."

Remove-Item build `
    -Recurse `
    -Force `
    -ErrorAction Ignore

Remove-Item dist `
    -Recurse `
    -Force `
    -ErrorAction Ignore


# --------------------------------------------
# 2. Build EXE
# --------------------------------------------

Write-Host ""
Write-Host "[2/5] Building EXE..."

pyinstaller `
    --clean `
    --noconfirm `
    EstimateTool.spec

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR: PyInstaller build failed."
    exit 1
}


# --------------------------------------------
# 3. Create distribution directory
# --------------------------------------------

Write-Host ""
Write-Host "[3/5] Creating distribution directory..."

New-Item `
    -ItemType Directory `
    -Path "dist\EstimateTool" `
    -Force |
    Out-Null


# --------------------------------------------
# 4. Copy distribution files
# --------------------------------------------

Write-Host ""
Write-Host "[4/5] Copying distribution files..."

# Move EXE
if (Test-Path "dist\EstimateTool.exe") {
    Move-Item `
        "dist\EstimateTool.exe" `
        "dist\EstimateTool\EstimateTool.exe" `
        -Force
}
else {
    Write-Host "ERROR: dist\EstimateTool.exe was not found."
    exit 1
}


# Copy resources
if (Test-Path "resources") {
    Copy-Item `
        "resources" `
        "dist\EstimateTool\resources" `
        -Recurse `
        -Force
}
else {
    Write-Host "WARNING: resources directory was not found."
}


# Copy .env
if (Test-Path ".env") {
    Copy-Item `
        ".env" `
        "dist\EstimateTool\.env" `
        -Force
}
else {
    Write-Host "WARNING: .env was not found."
}


# --------------------------------------------
# 5. Verify
# --------------------------------------------

Write-Host ""
Write-Host "[5/5] Verifying build result..."

$exePath = "dist\EstimateTool\EstimateTool.exe"

if (Test-Path $exePath) {

    Write-Host ""
    Write-Host "============================================"
    Write-Host "Build Success"
    Write-Host "============================================"
    Write-Host ""
    Write-Host "Output:"
    Write-Host "dist\EstimateTool"
    Write-Host ""

    Get-ChildItem `
        "dist\EstimateTool" `
        -Recurse
}
else {

    Write-Host ""
    Write-Host "ERROR: EstimateTool.exe was not found."
    exit 1
}