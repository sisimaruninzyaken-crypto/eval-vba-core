# import_to_excel.ps1
#
# VBA files -> Excel VBE auto-import script
#
# Workflow:
#   1. convert_encoding.ps1 -Reverse  (UTF-8 -> Shift-JIS)
#   2. Import all .bas / .cls / .frm into Excel VBE via COM
#   3. convert_encoding.ps1            (Shift-JIS -> UTF-8)
#
# Usage:
#   .\import_to_excel.ps1
#   .\import_to_excel.ps1 -ExcelPath "C:\path\to\book.xlsm"
#   .\import_to_excel.ps1 -KeepOpen   # do not close Excel after import
#
# Requirements:
#   Excel Trust Center > [x] Trust access to the VBA project object model
#

param(
    [string]$ExcelPath = "",
    [switch]$KeepOpen
)

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$scriptDir = $PSScriptRoot

# ---------------------------------------------------------------------------
# 0. Resolve Excel file path
# ---------------------------------------------------------------------------
if ($ExcelPath -eq "") {
    # Search for .xlsm files on Desktop (up to 2 levels deep), skip temp ~$ files
    $found = Get-ChildItem -Path "$env:USERPROFILE\Desktop" -Filter "*.xlsm" -Recurse -Depth 2 -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -notlike "~`$*" } |
             Select-Object -First 1
    if ($found) { $ExcelPath = $found.FullName }
}

if ($ExcelPath -eq "") {
    Write-Host "Excel file not found automatically."
    Write-Host "Please specify the path with -ExcelPath:"
    Write-Host "  .\import_to_excel.ps1 -ExcelPath `"C:\path\to\book.xlsm`""
    exit 1
}

Write-Host ""
Write-Host "=== VBA Import Script ==="
Write-Host "Script dir : $scriptDir"
Write-Host "Excel file : $ExcelPath"
Write-Host ""

# ---------------------------------------------------------------------------
# 1. UTF-8 -> Shift-JIS
# ---------------------------------------------------------------------------
Write-Host "[Step 1] Converting UTF-8 -> Shift-JIS..."
$convertScript = Join-Path $scriptDir "convert_encoding.ps1"
if (-not (Test-Path $convertScript)) {
    Write-Error "convert_encoding.ps1 not found in: $scriptDir"
    exit 1
}
& $convertScript -Reverse
Write-Host ""

# ---------------------------------------------------------------------------
# 2. Import into Excel VBE
# ---------------------------------------------------------------------------
Write-Host "[Step 2] Importing into Excel VBE..."

$excelWasOpen = $false
$excel = $null
$wb = $null

try {
    # Attach to running Excel if the target file is already open
    try {
        $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        foreach ($openWb in $excel.Workbooks) {
            if ($openWb.FullName -ieq $ExcelPath) {
                $wb = $openWb
                $excelWasOpen = $true
                Write-Host "  Attached to running Excel instance."
                break
            }
        }
    } catch {
        # Excel not running — will open below
    }

    if ($null -eq $excel) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
    }

    if ($null -eq $wb) {
        Write-Host "  Opening workbook..."
        $wb = $excel.Workbooks.Open($ExcelPath)
    }

    $vbp = $wb.VBProject

    # Check that VBA project access is enabled
    try {
        $null = $vbp.VBComponents.Count
    } catch {
        Write-Error @"
Cannot access VBProject. Enable it in Excel:
  File > Options > Trust Center > Trust Center Settings
  > Macro Settings > [x] Trust access to the VBA project object model
"@
        throw
    }

    # Collect files to import (bas, cls, frm — frx is handled automatically)
    $importFiles = @()
    foreach ($ext in @("*.bas", "*.cls", "*.frm")) {
        $importFiles += Get-ChildItem -Path $scriptDir -Filter $ext
    }

    Write-Host "  Found $($importFiles.Count) file(s) to import."

    $okCount = 0
    $skipCount = 0

    foreach ($f in $importFiles) {
        $name = $f.BaseName

        # Skip Excel's built-in ThisWorkbook / Sheet* class modules
        $builtinTypes = @("ThisWorkbook")
        $isBuiltin = $false
        foreach ($b in $builtinTypes) {
            if ($name -ieq $b) { $isBuiltin = $true; break }
        }
        # Also skip Sheet modules (Sheet1, Sheet2, ... or Japanese sheet names bound to sheets)
        try {
            $comp = $vbp.VBComponents.Item($name)
            # vbext_ct_Document = 100 (ThisWorkbook, Sheet modules)
            if ($comp.Type -eq 100) { $isBuiltin = $true }
        } catch { }

        if ($isBuiltin) {
            Write-Host "  [SKIP] $($f.Name)  (document module - use manual paste)"
            $skipCount++
            continue
        }

        # Remove existing component before re-import
        try {
            $existing = $vbp.VBComponents.Item($name)
            $vbp.VBComponents.Remove($existing)
        } catch {
            # Component does not exist yet - that's fine
        }

        $vbp.VBComponents.Import($f.FullName) | Out-Null
        Write-Host "  [OK]   $($f.Name)"
        $okCount++
    }

    Write-Host ""
    Write-Host "  Import complete: $okCount imported, $skipCount skipped."
    Write-Host ""

    # Save workbook
    Write-Host "  Saving workbook..."
    $wb.Save()
    Write-Host "  Saved."

} finally {
    # ---------------------------------------------------------------------------
    # 3. UTF-8 restore (runs even if import failed)
    # ---------------------------------------------------------------------------
    Write-Host ""
    Write-Host "[Step 3] Converting Shift-JIS -> UTF-8..."
    & $convertScript
    Write-Host ""

    # Close Excel only if we opened it
    if ($null -ne $excel) {
        if (-not $excelWasOpen -and -not $KeepOpen) {
            if ($null -ne $wb) { $wb.Close($false) }
            $excel.Quit()
            Write-Host "  Excel closed."
        } elseif ($KeepOpen) {
            $excel.Visible = $true
            Write-Host "  Excel left open (-KeepOpen)."
        } else {
            Write-Host "  Excel was already open - left as-is."
        }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

Write-Host "=== Done ==="
Write-Host ""
