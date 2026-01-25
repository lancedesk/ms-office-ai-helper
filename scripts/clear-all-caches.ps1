# Complete Office Add-in Cache Cleaner
# Run this script to completely remove all cached add-in data

Write-Host "Removing ALL Office add-in caches..." -ForegroundColor Cyan

# Close Word if running
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Seconds 2

# All known Office cache locations
$cachePaths = @(
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef",
    "$env:LOCALAPPDATA\Microsoft\Office\Wef", 
    "$env:APPDATA\Microsoft\Office\Wef",
    "$env:LOCALAPPDATA\Microsoft\Office\16.0\WebServiceCache",
    "$env:LOCALAPPDATA\Microsoft\Office\WebServiceCache",
    "$env:LOCALAPPDATA\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC",
    "$env:LOCALAPPDATA\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\LocalCache"
)

foreach ($p in $cachePaths) {
    if (Test-Path $p) {
        Remove-Item -Path "$p\*" -Recurse -Force -ErrorAction SilentlyContinue
        Write-Host "Cleared: $p" -ForegroundColor Green
    }
}

# Remove temp sideload documents
$tempDocs = Get-ChildItem -Path "$env:TEMP" -Filter "Word add-in*.docx" -ErrorAction SilentlyContinue
foreach ($doc in $tempDocs) {
    Remove-Item -Path $doc.FullName -Force -ErrorAction SilentlyContinue
    Write-Host "Removed temp doc: $($doc.Name)" -ForegroundColor Yellow
}

# Clear dist folder
$distPath = Join-Path $PSScriptRoot "dist"
if (Test-Path $distPath) {
    Remove-Item -Path "$distPath\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Cleared dist folder" -ForegroundColor Green
}

Write-Host ""
Write-Host "Done! All caches cleared." -ForegroundColor Cyan
Write-Host "Now run: npm start" -ForegroundColor White
