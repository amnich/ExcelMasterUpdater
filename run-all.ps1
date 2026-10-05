$ErrorActionPreference = 'Stop'
$tests = Get-ChildItem -Path "D:\Skrypty\ExcelMasterUpdater\Tests\Test-*.ps1"

Write-Host "Running all tests in Windows PowerShell 5.1..." -ForegroundColor Cyan
foreach ($t in $tests) {
    Write-Host "--- Running $($t.Name) [PS5] ---" -ForegroundColor Yellow
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $t.FullName
    if ($LASTEXITCODE -ne 0) { throw "Test $($t.Name) failed in PS5!" }
}

Write-Host "Running all tests in PowerShell 7+..." -ForegroundColor Cyan
foreach ($t in $tests) {
    Write-Host "--- Running $($t.Name) [PS7] ---" -ForegroundColor Yellow
    & pwsh.exe -NoProfile -File $t.FullName
    if ($LASTEXITCODE -ne 0) { throw "Test $($t.Name) failed in PS7!" }
}

Write-Host "All tests passed across both engines!" -ForegroundColor Green
