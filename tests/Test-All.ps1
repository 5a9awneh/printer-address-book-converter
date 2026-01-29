<#
.SYNOPSIS
    Comprehensive Automated Test Suite - All Scenarios

.DESCRIPTION
    Complete automated testing covering:
    1. Brand-pair conversions (16 combinations)
    2. Merge all samples to each target (4 merges)
    3. Convert merged outputs to other targets (16 combinations)
    4. Outlook mode (CLI)
    5. Cleanup and verification

.NOTES
    Run this single script for complete validation
#>

param(
    [switch]$KeepOutputs  # Keep generated files for inspection
)

$ErrorActionPreference = 'Continue'
$scriptPath = "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1"
$samplesDir = "$PSScriptRoot\demo_samples"

$stats = @{
    TotalTests = 0
    Passed = 0
    Failed = 0
    FailedTests = @()
}

function Write-TestHeader {
    param([string]$Title)
    Write-Host "`n================================================================" -ForegroundColor Cyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "================================================================`n" -ForegroundColor Cyan
}

function Test-ConversionSingle {
    param(
        [string]$SourceFile,
        [string]$TargetFile,
        [string]$TestName
    )
    
    $script:stats.TotalTests++
    Write-Host "  $TestName : " -NoNewline
    
    try {
        $output = & $scriptPath -SourcePath $SourceFile -TargetPath $TargetFile -NoInteractive *>&1
        $result = $output | Out-String
        
        if ($result -match 'Converted:\s*(\d+)' -or $result -match 'Writing\s+(\d+)\s+unique contacts') {
            $count = $matches[1]
            Write-Host "✅ $count contacts" -ForegroundColor Green
            $script:stats.Passed++
            return $true
        }
        else {
            Write-Host "❌ NO OUTPUT" -ForegroundColor Red
            $script:stats.Failed++
            $script:stats.FailedTests += $TestName
            return $false
        }
    }
    catch {
        Write-Host "❌ ERROR: $_" -ForegroundColor Red
        $script:stats.Failed++
        $script:stats.FailedTests += $TestName
        return $false
    }
}

function Test-MergeSingle {
    param(
        [string[]]$SourceFiles,
        [string]$TargetFile,
        [string]$TestName
    )
    
    $script:stats.TotalTests++
    Write-Host "  $TestName : " -NoNewline
    
    try {
        $output = & $scriptPath -Mode Merge -SourcePath $SourceFiles -TargetPath $TargetFile -NoInteractive *>&1
        $result = $output | Out-String
        
        if ($result -match 'After deduplication:\s*(\d+)') {
            $count = $matches[1]
            Write-Host "✅ $count unique" -ForegroundColor Green
            $script:stats.Passed++
            return $true
        }
        else {
            Write-Host "❌ NO OUTPUT" -ForegroundColor Red
            $script:stats.Failed++
            $script:stats.FailedTests += $TestName
            return $false
        }
    }
    catch {
        Write-Host "❌ ERROR: $_" -ForegroundColor Red
        $script:stats.Failed++
        $script:stats.FailedTests += $TestName
        return $false
    }
}

function Test-FieldPopulation {
    param(
        [string]$FilePath,
        [string]$Brand,
        [string]$TestName
    )
    
    $script:stats.TotalTests++
    Write-Host "  $TestName : " -NoNewline
    
    try {
        if (-not (Test-Path $FilePath)) {
            Write-Host "❌ FILE NOT FOUND" -ForegroundColor Red
            $script:stats.Failed++
            $script:stats.FailedTests += $TestName
            return $false
        }
        
        $content = Get-Content $FilePath -Raw
        $lines = Get-Content $FilePath
        
        # Find first data line (skip headers/comments/blanks)
        # Canon: # comments, blank line, header, data
        # Others: header, data
        $dataLine = $null
        $headerFound = $false
        foreach ($line in $lines) {
            # Skip comment/metadata/blank lines
            if ($line -match '^#' -or $line -match '^@Ver' -or $line.Trim() -eq '') { continue }
            
            # First non-comment, non-blank line is the header
            if (-not $headerFound) {
                $headerFound = $true
                continue
            }
            
            # Next non-blank line is data
            if ($line.Trim() -ne '') {
                $dataLine = $line
                break
            }
        }
        
        if (-not $dataLine) {
            Write-Host "❌ NO DATA ROWS" -ForegroundColor Red
            $script:stats.Failed++
            $script:stats.FailedTests += $TestName
            return $false
        }
        
        # Brand-specific validation
        $validated = $false
        switch ($Brand) {
            'Canon' {
                # Check for required Canon fields: objectclass, subdbid, protocol, indxid, uuid
                $validated = ($dataLine -match '"email"' -or $dataLine -match 'email,') -and
                             ($dataLine -match ',11,' -or $dataLine -match ',"11",') -and
                             ($dataLine -match 'smtp') -and
                             ($dataLine -match '[0-9a-f]{8}-[0-9a-f]{4}')
            }
            'Sharp' {
                # Check for required Sharp fields: address="data", search-id, category-id, frequently-used
                $validated = ($dataLine -match '"data"') -and
                             ($dataLine -match ',"[0-9]+",' -or $dataLine -match ',\d+,') -and
                             ($dataLine -match 'TRUE' -or $dataLine -match '"1"')
            }
            'Xerox' {
                # Check for required Xerox fields: XrxAddressBookId (sequential number), RefId (UUID format)
                $validated = ($dataLine -match '"\d+",' -or $dataLine -match '^\d+,') -and
                             ($dataLine -match '[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}')
            }
            'Develop' {
                # Check for required Develop fields: AbbrNo, Pinyin, SendMode, Level
                $validated = ($dataLine -match '"?\d+"?\t') -and
                             ($dataLine -match 'No') -and
                             ($dataLine -match 'Email') -and
                             ($dataLine -match 'Level')
            }
        }
        
        if ($validated) {
            Write-Host "✅ FIELDS OK" -ForegroundColor Green
            $script:stats.Passed++
            return $true
        }
        else {
            Write-Host "❌ MISSING FIELDS" -ForegroundColor Red
            $script:stats.Failed++
            $script:stats.FailedTests += $TestName
            return $false
        }
    }
    catch {
        Write-Host "❌ ERROR: $_" -ForegroundColor Red
        $script:stats.Failed++
        $script:stats.FailedTests += $TestName
        return $false
    }
}

# ================================================================
# SETUP
# ================================================================
Write-TestHeader "TEST SETUP"

Write-Host "Cleaning up previous test outputs..." -NoNewline
Get-ChildItem $samplesDir -Filter "*converted*.csv" -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem $samplesDir -Filter "Merged_*.csv" -ErrorAction SilentlyContinue | Remove-Item -Force
Write-Host " Done" -ForegroundColor Green

$testFiles = @{
    Canon   = "$samplesDir\Canon_Sample.csv"
    Sharp   = "$samplesDir\Sharp_Sample.csv"
    Xerox   = "$samplesDir\Xerox_Sample.csv"
    Develop = "$samplesDir\Develop_Sample.csv"
}

Write-Host "Verifying test files exist..." -NoNewline
$allExist = $true
foreach ($brand in $testFiles.Keys) {
    if (-not (Test-Path $testFiles[$brand])) {
        Write-Host "`n  ❌ Missing: $($testFiles[$brand])" -ForegroundColor Red
        $allExist = $false
    }
}

if (-not $allExist) {
    Write-Host "`nTest files missing. Exiting." -ForegroundColor Red
    exit 1
}
Write-Host " OK" -ForegroundColor Green

# ================================================================
# TEST SUITE 1: BRAND-PAIR CONVERSIONS (16 tests)
# ================================================================
Write-TestHeader "SUITE 1: Brand-Pair Conversions (16 combinations)"

$brands = @('Canon', 'Sharp', 'Xerox', 'Develop')
foreach ($source in $brands) {
    Write-Host "$source → All:" -ForegroundColor Yellow
    foreach ($target in $brands) {
        Test-ConversionSingle -SourceFile $testFiles[$source] -TargetFile $testFiles[$target] -TestName "$source → $target"
    }
}

# ================================================================
# TEST SUITE 2: MERGE ALL TO EACH TARGET (4 tests)
# ================================================================
Write-TestHeader "SUITE 2: Merge All 4 Samples → Each Target"

$allSources = @($testFiles.Canon, $testFiles.Sharp, $testFiles.Xerox, $testFiles.Develop)

foreach ($target in $brands) {
    Test-MergeSingle -SourceFiles $allSources -TargetFile $testFiles[$target] -TestName "ALL 4 → $target"
}

# ================================================================
# TEST SUITE 3: CONVERT MERGED OUTPUTS TO OTHER TARGETS (16 tests)
# ================================================================
Write-TestHeader "SUITE 3: Convert Merged Outputs → All Targets"

# Find the most recent merged files
Start-Sleep -Seconds 1
$mergedFiles = Get-ChildItem $samplesDir -Filter "Merged_converted_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 4

if ($mergedFiles.Count -lt 4) {
    Write-Host "⚠ Warning: Only $($mergedFiles.Count) merged files found, expected 4" -ForegroundColor Yellow
}

foreach ($mergedFile in $mergedFiles) {
    $sourceName = $mergedFile.Name
    Write-Host "$(Split-Path -Leaf $sourceName) → All:" -ForegroundColor Yellow
    
    foreach ($target in $brands) {
        Test-ConversionSingle -SourceFile $mergedFile.FullName -TargetFile $testFiles[$target] -TestName "Merged → $target"
    }
}

# ================================================================
# TEST SUITE 4: OUTLOOK MODE (4 tests - all targets)
# ================================================================
Write-TestHeader "SUITE 4: Outlook Import to All Targets"

$outlookTestFile = "$PSScriptRoot\temp_outlook_test.txt"
@"
SMITH, John <john.smith@example.com>;
DOE, Jane <jane.doe@example.com>;
JOHNSON, Bob <bob.johnson@example.com>;
"@ | Out-File -FilePath $outlookTestFile -Encoding UTF8

foreach ($target in $brands) {
    $script:stats.TotalTests++
    Write-Host "  Outlook → $target : " -NoNewline
    
    try {
        $output = & $scriptPath -Mode Outlook -SourcePath $outlookTestFile -TargetPath $testFiles[$target] -NoInteractive *>&1
        $result = $output | Out-String
        
        if ($result -match 'Writing\s+(\d+)\s+unique contacts') {
            $count = $matches[1]
            Write-Host "✅ $count contacts" -ForegroundColor Green
            $script:stats.Passed++
        }
        else {
            Write-Host "❌ NO OUTPUT" -ForegroundColor Red
            $script:stats.Failed++
            $script:stats.FailedTests += "Outlook → $target"
        }
    }
    catch {
        Write-Host "❌ ERROR: $_" -ForegroundColor Red
        $script:stats.Failed++
        $script:stats.FailedTests += "Outlook → $target"
    }
}

Remove-Item $outlookTestFile -Force -ErrorAction SilentlyContinue

# ================================================================
# TEST SUITE 5: FIELD POPULATION VALIDATION (8 tests)
# ================================================================
Write-TestHeader "SUITE 5: Field Population Validation"

# Test one conversion per brand to validate fields are populated
Write-Host "Generating test outputs for validation..." -ForegroundColor Yellow

# Canon → Sharp (validate Sharp fields)
& $scriptPath -SourcePath $testFiles.Canon -TargetPath $testFiles.Sharp -NoInteractive *>&1 | Out-Null
$sharpOutput = Get-ChildItem $samplesDir -Filter "*Canon*converted*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($sharpOutput) { Test-FieldPopulation -FilePath $sharpOutput.FullName -Brand 'Sharp' -TestName "Canon→Sharp Fields" }

# Sharp → Canon (validate Canon fields)
& $scriptPath -SourcePath $testFiles.Sharp -TargetPath $testFiles.Canon -NoInteractive *>&1 | Out-Null
$canonOutput = Get-ChildItem $samplesDir -Filter "*Sharp*converted*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($canonOutput) { Test-FieldPopulation -FilePath $canonOutput.FullName -Brand 'Canon' -TestName "Sharp→Canon Fields" }

# Xerox → Develop (validate Develop fields)
& $scriptPath -SourcePath $testFiles.Xerox -TargetPath $testFiles.Develop -NoInteractive *>&1 | Out-Null
$developOutput = Get-ChildItem $samplesDir -Filter "*Xerox*converted*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($developOutput) { Test-FieldPopulation -FilePath $developOutput.FullName -Brand 'Develop' -TestName "Xerox→Develop Fields" }

# Develop → Xerox (validate Xerox fields)
& $scriptPath -SourcePath $testFiles.Develop -TargetPath $testFiles.Xerox -NoInteractive *>&1 | Out-Null
$xeroxOutput = Get-ChildItem $samplesDir -Filter "*Develop*converted*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($xeroxOutput) { Test-FieldPopulation -FilePath $xeroxOutput.FullName -Brand 'Xerox' -TestName "Develop→Xerox Fields" }

# Outlook → All brands (validate field population for Outlook mode)
foreach ($target in $brands) {
    $outlookTestFile2 = "$PSScriptRoot\temp_outlook_validation.txt"
    "TEST, User <test@example.com>" | Out-File -FilePath $outlookTestFile2 -Encoding UTF8
    & $scriptPath -Mode Outlook -SourcePath $outlookTestFile2 -TargetPath $testFiles[$target] -NoInteractive *>&1 | Out-Null
    Remove-Item $outlookTestFile2 -Force -ErrorAction SilentlyContinue
    
    $outlookOutput = Get-ChildItem (Split-Path $samplesDir) -Filter "Outlook_converted_*.csv" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($outlookOutput) { 
        Test-FieldPopulation -FilePath $outlookOutput.FullName -Brand $target -TestName "Outlook→$target Fields"
        Remove-Item $outlookOutput.FullName -Force -ErrorAction SilentlyContinue
    }
}

# ================================================================
# CLEANUP
# ================================================================
Write-TestHeader "CLEANUP"

if (-not $KeepOutputs) {
    Write-Host "Removing generated test files..." -NoNewline
    $removed = 0
    Get-ChildItem $samplesDir -Filter "*converted*.csv" -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item $_.FullName -Force
        $removed++
    }
    Get-ChildItem $samplesDir -Filter "Merged_*.csv" -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item $_.FullName -Force
        $removed++
    }
    Get-ChildItem $PSScriptRoot -Filter "temp_outlook*.csv" -ErrorAction SilentlyContinue | ForEach-Object {
        Remove-Item $_.FullName -Force
        $removed++
    }
    Write-Host " $removed files removed" -ForegroundColor Green
}
else {
    Write-Host "Keeping output files for inspection (-KeepOutputs flag)" -ForegroundColor Yellow
}

# ================================================================
# SUMMARY
# ================================================================
Write-TestHeader "TEST SUMMARY"

Write-Host "  Total Tests:  $($stats.TotalTests)" -ForegroundColor White
Write-Host "  Passed:       $($stats.Passed)" -ForegroundColor Green
Write-Host "  Failed:       $($stats.Failed)" -ForegroundColor $(if ($stats.Failed -eq 0) { 'Green' } else { 'Red' })

$passRate = if ($stats.TotalTests -gt 0) { 
    [math]::Round(($stats.Passed / $stats.TotalTests) * 100, 1) 
} else { 0 }

Write-Host "  Pass Rate:    $passRate%" -ForegroundColor $(
    if ($passRate -ge 95) { 'Green' } 
    elseif ($passRate -ge 80) { 'Yellow' } 
    else { 'Red' }
)

if ($stats.Failed -gt 0) {
    Write-Host "`nFailed Tests:" -ForegroundColor Red
    foreach ($test in $stats.FailedTests) {
        Write-Host "  - $test" -ForegroundColor Red
    }
    Write-Host ""
    exit 1
}
else {
    Write-Host "`n✅ ALL TESTS PASSED!" -ForegroundColor Green
    Write-Host ""
    exit 0
}
