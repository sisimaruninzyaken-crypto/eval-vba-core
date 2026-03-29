# convert_encoding.ps1
#
# VBA export file encoding converter
# Export from Excel VBE -> Shift-JIS -> run this script -> UTF-8 (for Claude Code)
#
# Usage:
#   .\convert_encoding.ps1            # Shift-JIS -> UTF-8  (run after VBE export)
#   .\convert_encoding.ps1 -Reverse   # UTF-8 -> Shift-JIS  (run before VBE re-import if needed)
#
# Target: *.bas / *.cls / *.frm in the same directory as this script

param(
    [switch]$Reverse
)

$cp932 = [System.Text.Encoding]::GetEncoding(932)
$utf8  = New-Object System.Text.UTF8Encoding $false   # no BOM

if ($Reverse) {
    $from  = $utf8
    $to    = $cp932
    $label = "UTF-8 -> Shift-JIS"
} else {
    $from  = $cp932
    $to    = $utf8
    $label = "Shift-JIS -> UTF-8"
}

Write-Host "[$label] starting..."

$count = 0
foreach ($ext in @('*.bas', '*.cls', '*.frm')) {
    foreach ($f in Get-ChildItem -Path $PSScriptRoot -Filter $ext) {
        $bytes = [System.IO.File]::ReadAllBytes($f.FullName)
        $text  = $from.GetString($bytes)
        [System.IO.File]::WriteAllBytes($f.FullName, $to.GetBytes($text))
        Write-Host "  converted: $($f.Name)"
        $count++
    }
}

Write-Host "[$label] done -- $count file(s) converted."
