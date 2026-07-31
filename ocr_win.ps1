param([string]$Path, [string]$Lang = '')
$ErrorActionPreference = 'Stop'
[Windows.Media.Ocr.OcrEngine,Windows.Foundation,ContentType=WindowsRuntime]            | Out-Null
[Windows.Graphics.Imaging.BitmapDecoder,Windows.Foundation,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.StorageFile,Windows.Foundation,ContentType=WindowsRuntime]            | Out-Null
[Windows.Globalization.Language,Windows.Foundation,ContentType=WindowsRuntime]         | Out-Null

Add-Type -AssemblyName System.Runtime.WindowsRuntime
$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() |
  Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and
                 $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
function Await($op, $type) {
  $t = $asTask.MakeGenericMethod($type).Invoke($null, @($op))
  $t.Wait(-1) | Out-Null
  $t.Result
}

# --- pick the best engine: Spanish first (Mexican receipts), else English ---
$avail  = [Windows.Media.Ocr.OcrEngine]::AvailableRecognizerLanguages
$engine = $null
$order  = @()
if ($Lang) { $order += $Lang }
$order += @('es-MX','es-ES','es','en-US')
foreach ($tag in $order) {
  $hit = $avail | Where-Object { $_.LanguageTag -like "$tag*" } | Select-Object -First 1
  if ($hit) {
    $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromLanguage($hit)
    if ($engine) { break }
  }
}
if (-not $engine) { $engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages() }
if (-not $engine) { throw "no OCR engine available" }

$file    = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
$stream  = Await ($file.OpenAsync([Windows.Storage.FileAccessMode]::Read)) ([Windows.Storage.Streams.IRandomAccessStream])
$decoder = Await ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap  = Await ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$result  = Await ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])

[Console]::Error.WriteLine("engine=" + $engine.RecognizerLanguage.LanguageTag)
foreach ($line in $result.Lines) { [Console]::Out.WriteLine($line.Text) }
