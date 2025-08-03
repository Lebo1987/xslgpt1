# PowerShell script to sideload Excel Add-in
# Run this script as Administrator

Write-Host "Excel Add-in Sideload Helper" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Green

# Get the current directory
$currentDir = Get-Location
$manifestPath = Join-Path $currentDir "manifest-production.xml"

Write-Host "Looking for manifest file: $manifestPath" -ForegroundColor Yellow

if (Test-Path $manifestPath) {
    Write-Host "✅ Manifest file found!" -ForegroundColor Green
    Write-Host ""
    Write-Host "To sideload your add-in:" -ForegroundColor Cyan
    Write-Host "1. Open Excel" -ForegroundColor White
    Write-Host "2. Go to Developer tab (enable it in File → Options → Customize Ribbon)" -ForegroundColor White
    Write-Host "3. Click 'Excel Add-ins'" -ForegroundColor White
    Write-Host "4. Click 'Browse' and select: $manifestPath" -ForegroundColor White
    Write-Host "5. Click OK" -ForegroundColor White
    Write-Host ""
    Write-Host "Alternative method:" -ForegroundColor Cyan
    Write-Host "1. Open Excel" -ForegroundColor White
    Write-Host "2. Go to Insert → My Add-ins" -ForegroundColor White
    Write-Host "3. Look for 'Upload My Add-in' or 'Manage My Add-ins'" -ForegroundColor White
    Write-Host "4. Browse to: $manifestPath" -ForegroundColor White
    Write-Host ""
    Write-Host "Your Render URL: https://xslgpt1-excel-addin.onrender.com" -ForegroundColor Magenta
} else {
    Write-Host "❌ Manifest file not found!" -ForegroundColor Red
    Write-Host "Make sure you're running this script from the project directory." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 