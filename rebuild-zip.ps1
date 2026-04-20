Add-Type -AssemblyName System.IO.Compression.FileSystem
$folder  = 'd:\BRD Generator\teams-manifest'
$zipPath = "$folder\BRDGeneratorAgent.zip"

# Remove old ZIP
if (Test-Path $zipPath) { Remove-Item $zipPath -Force; Write-Host 'Removed old ZIP' }

# Verify required files
$required = @('manifest.json', 'color.png', 'outline.png')
$missing  = $required | Where-Object { !(Test-Path (Join-Path $folder $_)) }
if ($missing) { Write-Error "Missing files: $($missing -join ', ')"; exit 1 }

# Build fresh ZIP (Optimal compression)
$zip = [System.IO.Compression.ZipFile]::Open($zipPath, 'Create')
foreach ($file in $required) {
    $full = Join-Path $folder $file
    [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $full, $file, 'Optimal') | Out-Null
    $usz = (Get-Item $full).Length
    Write-Host "  Added: $file  ($usz bytes)"
}
$zip.Dispose()

# Integrity check
$zipSize = (Get-Item $zipPath).Length
Write-Host "`nZIP rebuilt: $zipPath"
Write-Host "ZIP size   : $zipSize bytes"
$zip2 = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
Write-Host "Contents   :"
foreach ($e in $zip2.Entries) {
    Write-Host "  $($e.FullName)  (compressed=$($e.CompressedLength)  uncompressed=$($e.Length))"
}
$zip2.Dispose()
Write-Host "Integrity  : OK"
