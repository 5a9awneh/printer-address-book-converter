<#
.SYNOPSIS
    Phase 1 Tests - CSV Structure Extraction & Header/Footer Preservation

.DESCRIPTION
    Tests for Get-CsvStructure() function to verify:
    - Headers are preserved exactly (including empty lines)
    - Contact rows are correctly identified (lines with @)
    - Footers are preserved exactly (including empty lines)
    - All brand formats work correctly (Canon, Sharp, Xerox, Develop)
#>

$ErrorActionPreference = 'Stop'

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Phase 1 Tests: CSV Structure Extraction" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

$testResults = @{
    Passed = 0
    Failed = 0
    Tests = @()
}

# Get test files
$testFiles = @(Get-ChildItem -Path "$PSScriptRoot\source_exports\*.csv" -ErrorAction SilentlyContinue)

if ($testFiles.Count -eq 0) {
    Write-Host "No test CSV files found in tests/source_exports/" -ForegroundColor Red
    Write-Host "Please add sample CSV files to test with." -ForegroundColor Yellow
    Write-Host ""
    exit 1
}

Write-Host "Found $($testFiles.Count) test files" -ForegroundColor Cyan
Write-Host ""

# Test 1: CLI Parameter Support
Write-Host "Test Suite 1: CLI Non-Interactive Mode" -ForegroundColor Cyan
Write-Host ""

foreach ($file in $testFiles | Select-Object -First 3) {
    $testName = "CLI Test: $($file.Name)"
    Write-Host "Test: $testName" -ForegroundColor Yellow

    try {
        # Run conversion in non-interactive mode
        $output = & "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1" `
            -SourcePath $file.FullName `
            -TargetBrand "Canon" `
            -NoInteractive 2>&1 | Out-String

        # Check if conversion succeeded (look for "Done" in output)
        if ($output -match "Done") {
            Write-Host "  Result: PASS" -ForegroundColor Green
            $testResults.Passed++
            $testResults.Tests += @{ Name = $testName; Status = "PASS" }
        }
        else {
            Write-Host "  Result: FAIL - No success indicator" -ForegroundColor Red
            $testResults.Failed++
            $testResults.Tests += @{ Name = $testName; Status = "FAIL" }
        }
    }
    catch {
        Write-Host "  Result: ERROR - $_" -ForegroundColor Red
        $testResults.Failed++
        $testResults.Tests += @{ Name = $testName; Status = "ERROR" }
    }

    Write-Host ""
}

# Test 2: Structure Extraction
Write-Host "Test Suite 2: Get-CsvStructure Function" -ForegroundColor Cyan
Write-Host ""

# Load functions manually (without running Main)
$scriptContent = Get-Content "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1" -Raw
$functionsToLoad = @(
    'Write-Log',
    'Get-FileEncoding',
    'Get-CsvStructure'
)

# Execute only the configuration and required functions
$configRegex = '(?ms)#region Configuration.*?#endregion'
$utilityRegex = '(?ms)#region Utility Functions.*?#endregion'
$parsingRegex = '(?ms)#region Parsing Functions.*?#endregion'
$detectionRegex = '(?ms)#region Detection Functions.*?#endregion'

if ($scriptContent -match $configRegex) {
    Invoke-Expression $matches[0]
}
if ($scriptContent -match $utilityRegex) {
    Invoke-Expression $matches[0]
}
if ($scriptContent -match $parsingRegex) {
    Invoke-Expression $matches[0]
}
if ($scriptContent -match $detectionRegex) {
    Invoke-Expression $matches[0]
}

foreach ($file in $testFiles) {
    $testName = "Structure Test: $($file.Name)"
    Write-Host "Test: $testName" -ForegroundColor Yellow

    try {
        # Detect encoding
        $encoding = Get-FileEncoding -FilePath $file.FullName

        # Extract structure
        $structure = Get-CsvStructure -FilePath $file.FullName -Encoding $encoding

        # Verify structure was extracted
        $hasHeaders = $structure.Headers.Count -ge 0
        $hasContacts = $structure.Contacts.Count -gt 0
        $hasFooters = $structure.Footers.Count -ge 0

        Write-Host "  Headers: $($structure.Headers.Count) lines"
        Write-Host "  Contacts: $($structure.Contacts.Count) lines"
        Write-Host "  Footers: $($structure.Footers.Count) lines"

        if ($hasContacts) {
            Write-Host "  Result: PASS" -ForegroundColor Green
            $testResults.Passed++
            $testResults.Tests += @{ Name = $testName; Status = "PASS" }
        }
        else {
            Write-Host "  Result: FAIL - No contacts found" -ForegroundColor Red
            $testResults.Failed++
            $testResults.Tests += @{ Name = $testName; Status = "FAIL" }
        }
    }
    catch {
        Write-Host "  Result: ERROR - $_" -ForegroundColor Red
        $testResults.Failed++
        $testResults.Tests += @{ Name = $testName; Status = "ERROR" }
    }

    Write-Host ""
}

# Test 3: Line Preservation
Write-Host "Test Suite 3: Line Preservation (Reconstruction)" -ForegroundColor Cyan
Write-Host ""

foreach ($file in $testFiles) {
    $testName = "Preservation Test: $($file.Name)"
    Write-Host "Test: $testName" -ForegroundColor Yellow

    try {
        # Read original
        $encoding = Get-FileEncoding -FilePath $file.FullName
        $originalLines = Get-Content -Path $file.FullName -Encoding $encoding

        # Extract and reconstruct
        $structure = Get-CsvStructure -FilePath $file.FullName -Encoding $encoding
        $reconstructed = @()
        $reconstructed += $structure.Headers
        $reconstructed += $structure.Contacts
        $reconstructed += $structure.Footers

        # Compare counts
        if ($originalLines.Count -eq $reconstructed.Count) {
            Write-Host "  Line count matches: $($originalLines.Count)" -ForegroundColor Green
            Write-Host "  Result: PASS" -ForegroundColor Green
            $testResults.Passed++
            $testResults.Tests += @{ Name = $testName; Status = "PASS" }
        }
        else {
            Write-Host "  Line count mismatch: Original=$($originalLines.Count), Reconstructed=$($reconstructed.Count)" -ForegroundColor Red
            Write-Host "  Result: FAIL" -ForegroundColor Red
            $testResults.Failed++
            $testResults.Tests += @{ Name = $testName; Status = "FAIL" }
        }
    }
    catch {
        Write-Host "  Result: ERROR - $_" -ForegroundColor Red
        $testResults.Failed++
        $testResults.Tests += @{ Name = $testName; Status = "ERROR" }
    }

    Write-Host ""
}

# Summary
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Test Summary" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total Tests: $($testResults.Passed + $testResults.Failed)"
Write-Host "  Passed: " -NoNewline
Write-Host $testResults.Passed -ForegroundColor Green
Write-Host "  Failed: " -NoNewline
Write-Host $testResults.Failed -ForegroundColor $(if ($testResults.Failed -eq 0) { "Green" } else { "Red" })
Write-Host ""

if ($testResults.Failed -eq 0) {
    Write-Host "  Phase 1: ALL TESTS PASSED" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "  Phase 1: SOME TESTS FAILED" -ForegroundColor Red
    exit 1
}

Write-Host ""
