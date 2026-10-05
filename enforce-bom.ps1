$files = @(
    'D:\Skrypty\ExcelMasterUpdater\Master-Updater.ps1',
    'D:\Skrypty\ExcelMasterUpdater\language.json',
    'D:\Skrypty\ExcelMasterUpdater\PROJECT_MEMORY.md',
    'D:\Skrypty\ExcelMasterUpdater\README.md'
)
foreach ($f in $files) {
    if (Test-Path $f) {
        $content = [System.IO.File]::ReadAllText($f)
        [System.IO.File]::WriteAllText($f, $content, [System.Text.UTF8Encoding]::new($true))
        Write-Host "Enforced UTF-8 BOM on $f"
    }
}
