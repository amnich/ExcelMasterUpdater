$errs = $null
$tokens = $null
[Management.Automation.Language.Parser]::ParseFile('D:\Skrypty\ExcelMasterUpdater\Master-Updater.ps1', [ref]$tokens, [ref]$errs)
if ($errs) {
    foreach ($e in $errs) {
        Write-Host "Error: $($e.Message) at line $($e.Extent.StartLineNumber) col $($e.Extent.StartColumnNumber)"
    }
    exit 1
} else {
    Write-Host "Syntax OK"
}
