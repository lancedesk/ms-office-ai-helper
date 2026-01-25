# Setup Loopback Exemption for Office Add-in Development
# Run this script ONCE as Administrator

Write-Host "Setting up localhost loopback exemption for Office Add-ins..." -ForegroundColor Cyan
Write-Host ""

# Add exemptions for WebView components
$packages = @(
    "Microsoft.Win32WebViewHost_cw5n1h2txyewy",
    "Microsoft.MicrosoftEdge_8wekyb3d8bbwe",
    "Microsoft.MicrosoftEdgeDevToolsClient_8wekyb3d8bbwe"
)

foreach ($pkg in $packages) {
    Write-Host "Adding loopback exemption for: $pkg" -ForegroundColor Yellow
    CheckNetIsolation LoopbackExempt -a -n="$pkg" 2>$null
}

Write-Host ""
Write-Host "Verifying exemptions..." -ForegroundColor Cyan
CheckNetIsolation LoopbackExempt -s

Write-Host ""
Write-Host "Done! You can now run 'npm start' without the Access Denied error." -ForegroundColor Green
Write-Host "Press any key to exit..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
