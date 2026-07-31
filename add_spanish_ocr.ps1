# Adds the Spanish (Mexico) OCR language so receipts read better.
# Needs Administrator - it will ask you for permission.

$ErrorActionPreference = 'Stop'

function Show-OcrLangs {
  [Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
  Write-Host ""
  Write-Host "  OCR languages now available:" -ForegroundColor Cyan
  foreach ($l in [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages) {
    Write-Host ("     " + $l.LanguageTag + "   " + $l.DisplayName)
  }
  Write-Host ""
}

# already there?
[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
$has = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages |
       Where-Object { $_.LanguageTag -like 'es*' }
if ($has) {
  Write-Host "  Spanish OCR is ALREADY installed - nothing to do." -ForegroundColor Green
  Show-OcrLangs
  Read-Host "  Press Enter to close"
  exit
}

# elevate if needed
$p = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Host "  This needs Administrator - asking Windows for permission..." -ForegroundColor Yellow
  Start-Process powershell -Verb RunAs -ArgumentList `
    "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
  exit
}

Write-Host ""
Write-Host "  Installing Spanish (Mexico) OCR..." -ForegroundColor Cyan
Write-Host "  This downloads a small package from Microsoft. Takes a minute or two."
Write-Host ""

$targets = @('Language.OCR~~~es-MX~0.0.1.0','Language.OCR~~~es-ES~0.0.1.0')
$done = $false
foreach ($t in $targets) {
  try {
    $cap = Get-WindowsCapability -Online -Name $t -ErrorAction Stop
    Write-Host ("  " + $cap.Name + "  ->  " + $cap.State)
    if ($cap.State -ne 'Installed') {
      Add-WindowsCapability -Online -Name $t -ErrorAction Stop | Out-Null
      Write-Host "  installed." -ForegroundColor Green
    } else {
      Write-Host "  already installed." -ForegroundColor Green
    }
    $done = $true
    break
  } catch {
    Write-Host ("  could not add " + $t + " : " + $_.Exception.Message) -ForegroundColor DarkYellow
  }
}

if (-not $done) {
  Write-Host ""
  Write-Host "  Automatic install did not work on this PC." -ForegroundColor Yellow
  Write-Host "  Do it by hand instead (2 minutes):"
  Write-Host "     1. Settings  >  Time & language  >  Language & region"
  Write-Host "     2. Add a language  ->  Spanish (Mexico)"
  Write-Host "     3. TICK 'Optical character recognition' in the options list"
  Write-Host "     4. Install, then run this again to check"
  Start-Process "ms-settings:regionlanguage"
}

Show-OcrLangs
Write-Host "  If es-MX is in the list above, receipts will now be read in Spanish."
Write-Host ""
Read-Host "  Press Enter to close"
