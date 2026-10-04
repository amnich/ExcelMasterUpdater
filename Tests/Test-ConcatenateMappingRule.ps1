# Test-ConcatenateMappingRule.ps1
# Unit and integration test for N:1 and 1:N multi-column concatenate mapping rules

$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Master-Updater.ps1"
. $scriptPath -NonInteractive

Write-Host "Master-Updater.ps1 loaded successfully for Test-ConcatenateMappingRule!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 1: 1 Incoming column -> 2 Base columns (Split with Concatenate & Separator)
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 1: 1 Incoming Column -> 2 Base Columns Projection ---" -ForegroundColor Cyan
$incomingData1 = @{
    'ImieNazwisko' = 'Zitokd Zicinski'
    'AdresPelny'   = 'Ziskoso 11/9, Gkiwicu'
}

$rules1 = @(
    @{
        BaseColumns   = @('ImieNazwisko')
        UpdateColumns = @('ImieNazwisko')
        MergeMode     = 'Exact'
    },
    @{
        BaseColumns   = @('AdresZamieszkania', 'Miasto')
        UpdateColumns = @('AdresPelny')
        MergeMode     = 'Concatenate'
        Separator     = ', '
    }
)

$proj1 = Get-ProjectedRow -IncomingValues $incomingData1 -MappingRules $rules1
if ($proj1['AdresZamieszkania'] -ne 'Ziskoso 11/9') {
    throw "Test 1 failed: Expected AdresZamieszkania to be 'Ziskoso 11/9', got '$($proj1['AdresZamieszkania'])'"
}
if ($proj1['Miasto'] -ne 'Gkiwicu') {
    throw "Test 1 failed: Expected Miasto to be 'Gkiwicu', got '$($proj1['Miasto'])'"
}
Write-Host "[PASS] Test 1: 1 Incoming -> 2 Base columns projected cleanly!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 2: 2 Incoming columns -> 1 Base column (Merge with Concatenate & Separator)
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 2: 2 Incoming Columns -> 1 Base Column Projection ---" -ForegroundColor Cyan
$incomingData2 = @{
    'ImieNazwisko' = 'Zitokd Zicinski'
    'UlicaDom'     = 'Ziskoso 11/9'
    'Miasto'       = 'Gkiwicu'
}

$rules2 = @(
    @{
        BaseColumns   = @('ImieNazwisko')
        UpdateColumns = @('ImieNazwisko')
        MergeMode     = 'Exact'
    },
    @{
        BaseColumns   = @('AdresPelny')
        UpdateColumns = @('UlicaDom', 'Miasto')
        MergeMode     = 'Concatenate'
        Separator     = ', '
    }
)

$proj2 = Get-ProjectedRow -IncomingValues $incomingData2 -MappingRules $rules2
if ($proj2['AdresPelny'] -ne 'Ziskoso 11/9, Gkiwicu') {
    throw "Test 2 failed: Expected AdresPelny to be 'Ziskoso 11/9, Gkiwicu', got '$($proj2['AdresPelny'])'"
}
Write-Host "[PASS] Test 2: 2 Incoming -> 1 Base column merged cleanly!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 3: Comparison Equality (2 Base Columns <-> 1 Incoming Column)
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 3: Comparison Equality (Unchanged Row) ---" -ForegroundColor Cyan
$baseRows3 = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values    = @{
            'ImieNazwisko'     = 'Zitokd Zicinski'
            'AdresZamieszkania'= 'Ziskoso 11/9'
            'Miasto'           = 'Gkiwicu'
        }
    }
)

$incomingRows3 = @(
    [PSCustomObject]@{
        SourceRowNumber = 2
        Values          = @{
            'ImieNazwisko' = 'Zitokd Zicinski'
            'AdresPelny'   = 'Ziskoso 11/9, Gkiwicu'
        }
    }
)

$res3 = @(Invoke-MasterCompare -BaseRows $baseRows3 -IncomingRows $incomingRows3 `
    -MappingRules $rules1 -BaseJoinKey @('ImieNazwisko') -IncomingJoinKey @('ImieNazwisko') `
    -BaseHeaders @('ImieNazwisko', 'AdresZamieszkania', 'Miasto'))

if ($res3.Count -ne 1) { throw "Test 3 failed: expected 1 review item, got $($res3.Count)" }
if ($res3[0].Status -ne 'Unchanged') {
    throw "Test 3 failed: expected row to be 'Unchanged', but got '$($res3[0].Status)' with changes: $($res3[0].Changes | Out-String)"
}
if ($res3[0].Changes.Count -ne 0) {
    throw "Test 3 failed: expected 0 changes, got $($res3[0].Changes.Count)"
}
Write-Host "[PASS] Test 3: Equal concatenated address identified as Unchanged (0 changes)!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 4: Comparison Difference (Incoming Address Changed)
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 4: Comparison Difference Detection ---" -ForegroundColor Cyan
$incomingRows4 = @(
    [PSCustomObject]@{
        SourceRowNumber = 2
        Values          = @{
            'ImieNazwisko' = 'Zitokd Zicinski'
            'AdresPelny'   = 'Nowa 10/2, Gkiwicu'
        }
    }
)

$res4 = @(Invoke-MasterCompare -BaseRows $baseRows3 -IncomingRows $incomingRows4 `
    -MappingRules $rules1 -BaseJoinKey @('ImieNazwisko') -IncomingJoinKey @('ImieNazwisko') `
    -BaseHeaders @('ImieNazwisko', 'AdresZamieszkania', 'Miasto'))

if ($res4.Count -ne 1) { throw "Test 4 failed: expected 1 review item, got $($res4.Count)" }
if ($res4[0].Status -ne 'Changed') {
    throw "Test 4 failed: expected row to be 'Changed', but got '$($res4[0].Status)'"
}
if ($res4[0].Changes.Count -ne 1) {
    throw "Test 4 failed: expected 1 cell change (AdresZamieszkania), but got $($res4[0].Changes.Count)"
}
$chg = $res4[0].Changes[0]
if ($chg.BaseColumn -ne 'AdresZamieszkania' -or $chg.OldValue -ne 'Ziskoso 11/9' -or $chg.NewValue -ne 'Nowa 10/2') {
    throw "Test 4 failed: unexpected change details: $($chg | Out-String)"
}
Write-Host "[PASS] Test 4: Genuine address change correctly detected (Old: 'Ziskoso 11/9' -> New: 'Nowa 10/2')!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 5: Real User Files Integration (If present)
# -----------------------------------------------------------------------------
$realBasePath = "C:\Users\adamm\Downloads\DDO_2\Wszystkie wnioski.xlsx"
$realIncPath  = "C:\Users\adamm\Downloads\DDO_2\lista_28.xlsx"

if ((Test-Path $realBasePath) -and (Test-Path $realIncPath)) {
    Write-Host "`n--- Test 5: Integration with real user files ---" -ForegroundColor Cyan
    $realBaseHdrs = [FastExcelHelper]::GetHeaders($realBasePath, $null)
    $realIncHdrs  = [FastExcelHelper]::GetHeaders($realIncPath, $null)

    $realBaseRowsRaw = [FastExcelHelper]::ReadSheet($realBasePath, $null)
    $realIncRowsRaw  = [FastExcelHelper]::ReadSheet($realIncPath, $null)

    $realRules = @(
        @{
            BaseColumns   = @('Imię i nazwisko ucznia/dziecka')
            UpdateColumns = @('Imię i nazwisko ucznia/dziecka')
            MergeMode     = 'Exact'
        },
        @{
            BaseColumns   = @('Adres zamieszkania dziecka i rodzica', 'Miasto Zamieszkania')
            UpdateColumns = @('Adres zamieszkania dziecka i rodzica')
            MergeMode     = 'Concatenate'
            Separator     = ', '
        }
    )

    $realComp = Invoke-MasterCompare -BaseRows $realBaseRowsRaw -IncomingRows $realIncRowsRaw `
        -MappingRules $realRules `
        -BaseJoinKey @('Imię i nazwisko ucznia/dziecka') `
        -IncomingJoinKey @('Imię i nazwisko ucznia/dziecka') `
        -BaseHeaders $realBaseHdrs

    $zicinskiItem = $realComp | Where-Object {
        $name = if ($_.IncomingRow.Values['Imię i nazwisko ucznia/dziecka']) { $_.IncomingRow.Values['Imię i nazwisko ucznia/dziecka'] } else { '' }
        $name -like '*Zici*'
    }

    if ($zicinskiItem) {
        Write-Host "Found Zitokd Zicinski in real comparison result! Status: $($zicinskiItem.Status)" -ForegroundColor Cyan
        $addrChanges = @($zicinskiItem.Changes | Where-Object { $_.BaseColumn -like '*Adres*' -or $_.BaseColumn -like '*Miasto*' })
        if ($addrChanges.Count -ne 0) {
            throw "Test 5 failed: Zitokd Zicinski has unexpected address changes: $($addrChanges | Out-String)"
        }
        Write-Host "[PASS] Test 5: Real user file record matched with zero address discrepancies!" -ForegroundColor Green
    } else {
        Write-Host "[SKIP] Test 5: Zitokd Zicinski not found in real files" -ForegroundColor Yellow
    }
}

Write-Host "`n==========================================================" -ForegroundColor Green
Write-Host " ALL CONCATENATE MAPPING RULE TESTS PASSED (100% SUCCESS)!" -ForegroundColor Green
Write-Host "==========================================================" -ForegroundColor Green
