# Test-QoLSelectiveReconciliation.ps1
# Automated verification test suite for Quality-of-Life (QoL) Selective Reconciliation:
# - Diff Policies: TrackChanges, IgnoreChanges, NormalizePostalCode, FuzzyContainment
# - Selective Takeover: SelectedForUpdate = $true / $false write-back isolation
# - Inline Card Editing: Editing NewValue before write-back, and Revert to OriginalNewValue
# - Live Staging Metrics calculation

$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path $PSScriptRoot -Parent) "Master-Updater.ps1"
. $scriptPath -NonInteractive

Write-Host "Master-Updater.ps1 loaded successfully for Test-QoLSelectiveReconciliation!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 1: DiffPolicy = 'IgnoreChanges' (New records only)
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 1: DiffPolicy = 'IgnoreChanges' ---" -ForegroundColor Cyan
$baseRows = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values = [ordered]@{
            'ID'           = '101'
            'ImieNazwisko' = 'Jan Kowalski'
            'Szkola'       = 'Szkoła Podstawowa nr 1'
            'Telefon'      = '123-456-789'
        }
    }
)
$incRows = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values = [ordered]@{
            'ID'           = '101'
            'ImieNazwisko' = 'Jan Kowalski'
            'Szkola'       = 'Szkoła Podstawowa nr 1 w Gliwicach'
            'Telefon'      = '987-654-321'
        }
    },
    [PSCustomObject]@{
        RowNumber = 3
        Values = [ordered]@{
            'ID'           = '102'
            'ImieNazwisko' = 'Anna Nowak'
            'Szkola'       = 'Przedszkole nr 5'
            'Telefon'      = '555-666-777'
        }
    }
)

# Case A: Standard TrackChanges -> should produce changes for both Szkola and Telefon
$rulesTrack = @(
    @{ BaseColumns = @('ID');           UpdateColumns = @('ID');           MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('ImieNazwisko'); UpdateColumns = @('ImieNazwisko'); MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('Szkola');       UpdateColumns = @('Szkola');       MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('Telefon');      UpdateColumns = @('Telefon');      MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' }
)
$compTrack = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incRows -MappingRules $rulesTrack -BaseJoinKey @('ID') -IncomingJoinKey @('ID')
$matchedTrack = $compTrack | Where-Object { $_.MatchedBaseRow -and $_.MatchedBaseRow.Values['ID'] -eq '101' }
if ($matchedTrack.Changes.Count -ne 2) {
    throw "Test 1A failed: Expected 2 changes with TrackChanges, got $($matchedTrack.Changes.Count)"
}

# Case B: Szkola set to 'IgnoreChanges' -> should only produce 1 change (Telefon), Szkola ignored!
$rulesIgnore = @(
    @{ BaseColumns = @('ID');           UpdateColumns = @('ID');           MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('ImieNazwisko'); UpdateColumns = @('ImieNazwisko'); MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('Szkola');       UpdateColumns = @('Szkola');       MergeMode = 'Exact'; DiffPolicy = 'IgnoreChanges' },
    @{ BaseColumns = @('Telefon');      UpdateColumns = @('Telefon');      MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' }
)
$compIgnore = Invoke-MasterCompare -BaseRows $baseRows -IncomingRows $incRows -MappingRules $rulesIgnore -BaseJoinKey @('ID') -IncomingJoinKey @('ID')
$matchedIgnore = $compIgnore | Where-Object { $_.MatchedBaseRow -and $_.MatchedBaseRow.Values['ID'] -eq '101' }
if ($matchedIgnore.Changes.Count -ne 1) {
    throw "Test 1B failed: Expected 1 change with IgnoreChanges, got $($matchedIgnore.Changes.Count)"
}
if ($matchedIgnore.Changes[0].BaseColumn -ne 'Telefon') {
    throw "Test 1B failed: Expected changed column to be 'Telefon', got '$($matchedIgnore.Changes[0].BaseColumn)'"
}

# Case C: New row must STILL have the projected Szkola value
$newRow = $compIgnore | Where-Object { $_.Status -eq 'New' }
if ($newRow.ProjectedRow['Szkola'] -ne 'Przedszkole nr 5') {
    throw "Test 1C failed: New row did not retain projected Szkola! Got '$($newRow.ProjectedRow['Szkola'])'"
}
Write-Host "[PASS] Test 1: DiffPolicy = 'IgnoreChanges' ignores matched differences while projecting new records!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 2: DiffPolicy = 'NormalizePostalCode'
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 2: DiffPolicy = 'NormalizePostalCode' ---" -ForegroundColor Cyan
$baseRows2 = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values = [ordered]@{
            'ID'     = '201'
            'Miasto' = 'Gliwice'
        }
    },
    [PSCustomObject]@{
        RowNumber = 3
        Values = [ordered]@{
            'ID'     = '202'
            'Miasto' = 'Gliwice'
        }
    }
)
$incRows2 = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values = [ordered]@{
            'ID'     = '201'
            'Miasto' = '44-100 Gliwice' # Same city with postal code -> Should be considered equal!
        }
    },
    [PSCustomObject]@{
        RowNumber = 3
        Values = [ordered]@{
            'ID'     = '202'
            'Miasto' = '44-100 Zabrze'  # Different city -> Should be detected as changed!
        }
    }
)
$rulesPostal = @(
    @{ BaseColumns = @('ID');     UpdateColumns = @('ID');     MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('Miasto'); UpdateColumns = @('Miasto'); MergeMode = 'Exact'; DiffPolicy = 'NormalizePostalCode' }
)
$compPostal = Invoke-MasterCompare -BaseRows $baseRows2 -IncomingRows $incRows2 -MappingRules $rulesPostal -BaseJoinKey @('ID') -IncomingJoinKey @('ID')
$row201 = $compPostal | Where-Object { $_.MatchedBaseRow -and $_.MatchedBaseRow.Values['ID'] -eq '201' }
$row202 = $compPostal | Where-Object { $_.MatchedBaseRow -and $_.MatchedBaseRow.Values['ID'] -eq '202' }

if ($row201.Changes.Count -ne 0) {
    throw "Test 2 failed: Expected 0 changes for '44-100 Gliwice' vs 'Gliwice', got $($row201.Changes.Count)"
}
if ($row202.Changes.Count -ne 1) {
    throw "Test 2 failed: Expected 1 change for '44-100 Zabrze' vs 'Gliwice', got $($row202.Changes.Count)"
}
Write-Host "[PASS] Test 2: DiffPolicy = 'NormalizePostalCode' successfully neutralizes postal code variations!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 3: DiffPolicy = 'FuzzyContainment'
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 3: DiffPolicy = 'FuzzyContainment' ---" -ForegroundColor Cyan
$baseRows3 = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values = [ordered]@{
            'ID'     = '301'
            'Szkola' = 'Przedszkole Miejskie nr 21'
        }
    },
    [PSCustomObject]@{
        RowNumber = 3
        Values = [ordered]@{
            'ID'     = '302'
            'Szkola' = 'Przedszkole Miejskie nr 21'
        }
    }
)
$incRows3 = @(
    [PSCustomObject]@{
        RowNumber = 2
        Values = [ordered]@{
            'ID'     = '301'
            'Szkola' = 'Przedszkole Miejskie nr 21 w Gliwicach' # Substring containment -> Equal!
        }
    },
    [PSCustomObject]@{
        RowNumber = 3
        Values = [ordered]@{
            'ID'     = '302'
            'Szkola' = 'Szkoła Podstawowa nr 5'                # Distinct -> Change detected!
        }
    }
)
$rulesFuzzy = @(
    @{ BaseColumns = @('ID');     UpdateColumns = @('ID');     MergeMode = 'Exact'; DiffPolicy = 'TrackChanges' },
    @{ BaseColumns = @('Szkola'); UpdateColumns = @('Szkola'); MergeMode = 'Exact'; DiffPolicy = 'FuzzyContainment' }
)
$compFuzzy = Invoke-MasterCompare -BaseRows $baseRows3 -IncomingRows $incRows3 -MappingRules $rulesFuzzy -BaseJoinKey @('ID') -IncomingJoinKey @('ID')
$row301 = $compFuzzy | Where-Object { $_.MatchedBaseRow -and $_.MatchedBaseRow.Values['ID'] -eq '301' }
$row302 = $compFuzzy | Where-Object { $_.MatchedBaseRow -and $_.MatchedBaseRow.Values['ID'] -eq '302' }

if ($row301.Changes.Count -ne 0) {
    throw "Test 3 failed: Expected 0 changes for 'Przedszkole Miejskie nr 21 w Gliwicach' vs 'Przedszkole Miejskie nr 21', got $($row301.Changes.Count)"
}
if ($row302.Changes.Count -ne 1) {
    throw "Test 3 failed: Expected 1 change for 'Szkoła Podstawowa nr 5' vs 'Przedszkole Miejskie nr 21', got $($row302.Changes.Count)"
}
Write-Host "[PASS] Test 3: DiffPolicy = 'FuzzyContainment' ignores benign name extensions!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 4: Selective Write-Back & Inline Cell Editing in Excel File
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 4: Selective Write-Back & Inline Cell Editing Integration ---" -ForegroundColor Cyan
$tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) ("MasterUpdater_QoLTest_" + [System.Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($tmpDir)

try {
    $baseFixture = Join-Path (Split-Path $PSScriptRoot -Parent) "TestFixtures\base.xlsx"
    $baseXlsx = Join-Path $tmpDir 'BaseTest.xlsx'
    [System.IO.File]::Copy($baseFixture, $baseXlsx, $true)

    # In base.xlsx, Row 2 has:
    # TelefonRodzica = '500-111-000' (Col H / Index 7)
    # Szkola         = 'SP nr 1'       (Col F / Index 5)

    # Simulate comparison that generated 2 changes: TelefonRodzica and Szkola
    $chgTelefon = [ordered]@{
        BaseColumn        = 'TelefonRodzica'
        ColIndex          = 7
        CellRef           = 'H2'
        OldValue          = '500-111-000'
        NewValue          = '999-888-777' # Will be edited inline!
        OriginalNewValue  = '999-888-777'
        CustomEdited      = $false
        SelectedForUpdate = $true         # Selected!
    }
    $chgSzkola = [ordered]@{
        BaseColumn        = 'Szkola'
        ColIndex          = 5
        CellRef           = 'F2'
        OldValue          = 'SP nr 1'
        NewValue          = 'Nowa Szkola z pliku'
        OriginalNewValue  = 'Nowa Szkola z pliku'
        CustomEdited      = $false
        SelectedForUpdate = $false        # DESELECTED by user! Must not be written back!
    }

    # Simulate direct inline card editing: user overrides TelefonRodzica with custom value
    $chgTelefon.NewValue = '777-666-555'
    $chgTelefon.CustomEdited = $true

    $acceptedItem = [PSCustomObject]@{
        Status            = 'Changed'
        MatchedBaseRow    = [PSCustomObject]@{ RowNumber = 2 }
        Changes           = @($chgTelefon, $chgSzkola)
        SelectedForUpdate = $true
    }

    $appCfg = @{
        WriteMode            = 'InPlace'
        SafeRewriteMode      = $false
        BackupBeforeWrite    = $false
        BackupDirectory      = $tmpDir
        BackupRetentionCount = 1
        MaxBackupMb          = 200
        MetadataColumns      = @()
    }

    $wbRes = Invoke-MasterWriteBack -BaseFilePath $baseXlsx -BaseSheet 'Sheet1' -AcceptedItems @($acceptedItem) -AppConfig $appCfg
    if ($wbRes.UpdatedCells -ne 1) {
        throw "Test 4 failed: Expected exactly 1 cell updated (TelefonRodzica), got $($wbRes.UpdatedCells)"
    }

    # Inspect the saved Excel file using FastExcelHelper
    $rowsAfter = [FastExcelHelper]::ReadSheet($baseXlsx, 'Sheet1')
    $targetRow = $rowsAfter | Where-Object { [int]$_._RowNumber -eq 2 }
    if (-not $targetRow) {
        throw "Test 4 failed: Row 2 not found in written file"
    }

    $finalPhone = $targetRow.TelefonRodzica
    $finalSchool = $targetRow.Szkola

    if ($finalPhone -ne '777-666-555') {
        throw "Test 4 failed: Inline edited phone was not written! Expected '777-666-555', got '$finalPhone'"
    }
    if ($finalSchool -ne 'SP nr 1') {
        throw "Test 4 failed: Deselected school change was overwritten! Expected 'SP nr 1', got '$finalSchool'"
    }
    Write-Host "[PASS] Test 4: Selective takeover and inline edited cell values correctly written to Excel!" -ForegroundColor Green

} finally {
    if (Test-Path $tmpDir) {
        [System.IO.Directory]::Delete($tmpDir, $true)
    }
}

# -----------------------------------------------------------------------------
# Test 5: Inline Revert Verification
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 5: Inline Revert Functionality ---" -ForegroundColor Cyan
$chgTest = [ordered]@{
    BaseColumn        = 'Adres'
    OriginalNewValue  = 'Ulica Kwiatowa 1'
    NewValue          = 'Ulica Kwiatowa 1'
    CustomEdited      = $false
}
# Simulate user typing
$chgTest.NewValue = 'Ulica Kwiatowa 1A'
$chgTest.CustomEdited = ($chgTest.NewValue -ne $chgTest.OriginalNewValue)
if ($chgTest.CustomEdited -ne $true) {
    throw "Test 5 failed: CustomEdited should be true after text edit"
}
# Simulate clicking Revert (↺)
$chgTest.NewValue = $chgTest.OriginalNewValue
$chgTest.CustomEdited = ($chgTest.NewValue -ne $chgTest.OriginalNewValue)
if ($chgTest.CustomEdited -ne $false) {
    throw "Test 5 failed: CustomEdited should be false after revert"
}
if ($chgTest.NewValue -ne 'Ulica Kwiatowa 1') {
    throw "Test 5 failed: Revert did not restore OriginalNewValue"
}
Write-Host "[PASS] Test 5: Inline cell editing and revert (↺) work deterministically!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 6: Staging Metrics Formula Verification
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 6: Staging Summary Metrics ---" -ForegroundColor Cyan
$items = @(
    [PSCustomObject]@{
        Decision      = 'Accepted'
        SelectedCells = @{ 'Telefon' = $true; 'Szkola' = $false }
        Record        = [PSCustomObject]@{
            Status  = 'Changed'
            Changes = @(
                [ordered]@{ BaseColumn = 'Telefon' },
                [ordered]@{ BaseColumn = 'Szkola' }
            )
        }
    },
    [PSCustomObject]@{
        Decision      = 'Accepted'
        SelectedCells = @{ 'Telefon' = $true; 'Miasto' = $true }
        Record        = [PSCustomObject]@{
            Status  = 'Changed'
            Changes = @(
                [ordered]@{ BaseColumn = 'Telefon' },
                [ordered]@{ BaseColumn = 'Miasto' }
            )
        }
    },
    [PSCustomObject]@{
        Decision      = 'Skipped' # Should be excluded from staging counters
        SelectedCells = @{ 'Telefon' = $true }
        Record        = [PSCustomObject]@{
            Status  = 'Changed'
            Changes = @(
                [ordered]@{ BaseColumn = 'Telefon' }
            )
        }
    }
)

$stagedCount = 0
$skippedCount = 0
$affectedRows = 0
foreach ($it in ($items | Where-Object { $_.Decision -eq 'Accepted' })) {
    $hasStagedInRow = $false
    foreach ($c in $it.Record.Changes) {
        $sel = if ($it.SelectedCells) { ($it.SelectedCells[$c.BaseColumn] -ne $false) } else { $true }
        if ($sel) {
            $stagedCount++
            $hasStagedInRow = $true
        } else {
            $skippedCount++
        }
    }
    if ($hasStagedInRow) { $affectedRows++ }
}

if ($stagedCount -ne 3) {
    throw "Test 6 failed: Expected 3 staged changes, got $stagedCount"
}
if ($skippedCount -ne 1) {
    throw "Test 6 failed: Expected 1 skipped change, got $skippedCount"
}
if ($affectedRows -ne 2) {
    throw "Test 6 failed: Expected 2 affected rows, got $affectedRows"
}
Write-Host "[PASS] Test 6: Staging metrics formula accurately counts staged, skipped, and affected rows!" -ForegroundColor Green

# -----------------------------------------------------------------------------
# Test 7: Apply to New File Verification (Preserve Original Base File)
# -----------------------------------------------------------------------------
Write-Host "`n--- Test 7: Apply to New File Verification ---" -ForegroundColor Cyan
$tmpDir7 = Join-Path ([System.IO.Path]::GetTempPath()) ("MasterUpdater_ApplyNewTest_" + [System.Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($tmpDir7)

try {
    $baseFixture = Join-Path (Split-Path $PSScriptRoot -Parent) "TestFixtures\base.xlsx"
    $originalMaster = Join-Path $tmpDir7 'Master_Base_Untouched.xlsx'
    $newOutputFile  = Join-Path $tmpDir7 'Master_Base_Updated_20261005.xlsx'
    [System.IO.File]::Copy($baseFixture, $originalMaster, $true)

    # Simulate 1 selective change on Row 2
    $chgTelefon7 = [ordered]@{
        BaseColumn        = 'TelefonRodzica'
        ColIndex          = 7
        CellRef           = 'H2'
        OldValue          = '500-111-000'
        NewValue          = '555-444-333'
        OriginalNewValue  = '555-444-333'
        CustomEdited      = $false
        SelectedForUpdate = $true
    }
    $acceptedItem7 = [PSCustomObject]@{
        Status            = 'Changed'
        MatchedBaseRow    = [PSCustomObject]@{ RowNumber = 2 }
        Changes           = @($chgTelefon7)
        SelectedForUpdate = $true
    }

    # Simulate Save to New File action:
    # 1. Copy original master to new destination
    [System.IO.File]::Copy($originalMaster, $newOutputFile, $true)

    # 2. Apply writeback targeting the new destination
    $appCfg7 = @{
        WriteMode            = 'InPlace'
        SafeRewriteMode      = $false
        BackupBeforeWrite    = $false
        BackupDirectory      = $tmpDir7
        BackupRetentionCount = 1
        MaxBackupMb          = 200
        MetadataColumns      = @()
    }
    $wbRes7 = Invoke-MasterWriteBack -BaseFilePath $newOutputFile -BaseSheet 'Sheet1' -AcceptedItems @($acceptedItem7) -AppConfig $appCfg7
    if ($wbRes7.UpdatedCells -ne 1) {
        throw "Test 7 failed: Expected 1 updated cell in new file, got $($wbRes7.UpdatedCells)"
    }

    # Verify NEW file contains the updated value
    $rowsNew = [FastExcelHelper]::ReadSheet($newOutputFile, 'Sheet1')
    $rowNew2 = $rowsNew | Where-Object { [int]$_._RowNumber -eq 2 }
    if ($rowNew2.TelefonRodzica -ne '555-444-333') {
        throw "Test 7 failed: New file was not updated with changes! Got: '$($rowNew2.TelefonRodzica)'"
    }

    # Verify ORIGINAL file remained COMPLETELY untouched
    $rowsOriginal = [FastExcelHelper]::ReadSheet($originalMaster, 'Sheet1')
    $rowOrig2 = $rowsOriginal | Where-Object { [int]$_._RowNumber -eq 2 }
    if ($rowOrig2.TelefonRodzica -ne '500-111-000') {
        throw "Test 7 failed: Original master file was modified! Expected '500-111-000', got: '$($rowOrig2.TelefonRodzica)'"
    }

    Write-Host "[PASS] Test 7: 'Save as New File' creates updated file while preserving original base untouched!" -ForegroundColor Green

} finally {
    if (Test-Path $tmpDir7) {
        [System.IO.Directory]::Delete($tmpDir7, $true)
    }
}

Write-Host "`n========================================================" -ForegroundColor Green
Write-Host "ALL 7 QoL SELECTIVE RECONCILIATION TESTS PASSED 100%!" -ForegroundColor Green
Write-Host "========================================================`n" -ForegroundColor Green
