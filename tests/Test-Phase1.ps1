<#
.SYNOPSIS
    Phase 1 Tests - CSV Structure Extraction & Header/Footer Preservation

.DESCRIPTION
    Tests for Extract-CsvStructure() function to verify:
    - Headers are preserved exactly (including empty lines)
    - Contact rows are correctly identified (lines with @)
    - Footers are preserved exactly (including empty lines)
    - All brand formats work correctly (Canon, Sharp, Xerox, Develop)
#>

# Load only the functions we need (not Main)
$scriptContent = Get-Content "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1" -Raw
$functionsOnly = $scriptContent -replace '(?ms)^function Main \{.*?^#endregion.*$', ''
Invoke-Expression $functionsOnly

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

function Test-CsvStructure {
    param(
        [string]$TestName,
        [string]$FilePath,
        [int]$ExpectedHeaders,
        [int]$ExpectedContacts,
        [int]$ExpectedFooters
    )

    Write-Host "Test: $TestName" -ForegroundColor Yellow
    Write-Host "  File: $(Split-Path -Leaf $FilePath)"

    try {
        if (-not (Test-Path $FilePath)) {
            Write-Host "  SKIP - File not found" -ForegroundColor Gray
            return
        }

        # Detect encoding
        $encoding = Detect-Encoding -FilePath $FilePath

        # Extract structure
        $structure = Extract-CsvStructure -FilePath $FilePath -Encoding $encoding

        # Verify counts
        $headerMatch = $structure.Headers.Count -eq $ExpectedHeaders
        $contactMatch = $structure.Contacts.Count -eq $ExpectedContacts
        $footerMatch = $structure.Footers.Count -eq $ExpectedFooters

        Write-Host "  Headers: $($structure.Headers.Count) (expected $ExpectedHeaders) " -NoNewline
        if ($headerMatch) { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red }

        Write-Host "  Contacts: $($structure.Contacts.Count) (expected $ExpectedContacts) " -NoNewline
        if ($contactMatch) { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red }

        Write-Host "  Footers: $($structure.Footers.Count) (expected $ExpectedFooters) " -NoNewline
        if ($footerMatch) { Write-Host "PASS" -ForegroundColor Green } else { Write-Host "FAIL" -ForegroundColor Red }

        if ($headerMatch -and $contactMatch -and $footerMatch) {
            Write-Host "  Result: PASS" -ForegroundColor Green
            $testResults.Passed++
            $testResults.Tests += @{ Name = $TestName; Status = "PASS" }
        }
        else {
            Write-Host "  Result: FAIL" -ForegroundColor Red
            $testResults.Failed++
            $testResults.Tests += @{ Name = $TestName; Status = "FAIL" }
        }
    }
    catch {
        Write-Host "  Result: ERROR - $_" -ForegroundColor Red
        $testResults.Failed++
        $testResults.Tests += @{ Name = $TestName; Status = "ERROR" }
    }

    Write-Host ""
}

function Test-LinePreservation {
    param(
        [string]$TestName,
        [string]$FilePath
    )

    Write-Host "Test: $TestName" -ForegroundColor Yellow
    Write-Host "  File: $(Split-Path -Leaf $FilePath)"

    try {
        if (-not (Test-Path $FilePath)) {
            Write-Host "  SKIP - File not found" -ForegroundColor Gray
            return
        }

        # Read original file
        $encoding = Detect-Encoding -FilePath $FilePath
        $originalLines = Get-Content -Path $FilePath -Encoding $encoding

        # Extract and reconstruct
        $structure = Extract-CsvStructure -FilePath $FilePath -Encoding $encoding
        $reconstructed = @()
        $reconstructed += $structure.Headers
        $reconstructed += $structure.Contacts
        $reconstructed += $structure.Footers

        # Compare line counts
        if ($originalLines.Count -ne $reconstructed.Count) {
            Write-Host "  Line count mismatch: Original=$($originalLines.Count), Reconstructed=$($reconstructed.Count)" -ForegroundColor Red
            Write-Host "  Result: FAIL" -ForegroundColor Red
            $testResults.Failed++
            $testResults.Tests += @{ Name = $TestName; Status = "FAIL" }
        }
        else {
            # Check if all lines match
            $allMatch = $true
            for ($i = 0; $i -lt $originalLines.Count; $i++) {
                if ($originalLines[$i] -ne $reconstructed[$i]) {
                    Write-Host "  Line $($i+1) mismatch" -ForegroundColor Red
                    Write-Host "    Original:      '$($originalLines[$i])'" -ForegroundColor Gray
                    Write-Host "    Reconstructed: '$($reconstructed[$i])'" -ForegroundColor Gray
                    $allMatch = $false
                    break
                }
            }

            if ($allMatch) {
                Write-Host "  All $($originalLines.Count) lines match exactly" -NoNewline
                Write-Host " " -NoNewline
                Write-Host "Pass" -ForegroundColor Green
                Write-Host "  Result: PASS" -ForegroundColor Green
                $testResults.Passed++
                $testResults.Tests += @{ Name = $TestName; Status = "PASS" }
            }
            else {
                Write-Host "  Result: FAIL" -ForegroundColor Red
                $testResults.Failed++
                $testResults.Tests += @{ Name = $TestName; Status = "FAIL" }
            }
        }
    }
    catch {
        Write-Host "  Result: ERROR - $_" -ForegroundColor Red
        $testResults.Failed++
        $testResults.Tests += @{ Name = $TestName; Status = "ERROR" }
    }

    Write-Host ""
}

# Run tests on sample files
$testFiles = @(
    @{
        Name = "Canon Sample"
        Path = "$PSScriptRoot\source_exports\10.52.18.18_Canon_imageFORCE 6160.csv"
        ExpectedHeaders = 4
        ExpectedContacts = 24
        ExpectedFooters = 0
    },
    @{
        Name = "Sharp Sample"
        Path = "$PSScriptRoot\source_exports\10.52.18.61_SHARP_MX-3051.csv"
        ExpectedHeaders = 1
        ExpectedContacts = 13
        ExpectedFooters = 0
    },
    @{
        Name = "Xerox Sample"
        Path = "$PSScriptRoot\source_exports\10.52.18.45_Xerox AltaLink B8170.csv"
        ExpectedHeaders = 1
        ExpectedContacts = 27
        ExpectedFooters = 0
    },
    @{
        Name = "Develop Sample"
        Path = "$PSScriptRoot\source_exports\10.52.18.40_Develop ineo+ 224e.csv"
        ExpectedHeaders = 2
        ExpectedContacts = 15
        ExpectedFooters = 0
    }
)

# Test structure extraction
Write-Host "Test Suite 1: Structure Extraction" -ForegroundColor Cyan
Write-Host ""

foreach ($test in $testFiles) {
    Test-CsvStructure -TestName $test.Name `
                      -FilePath $test.Path `
                      -ExpectedHeaders $test.ExpectedHeaders `
                      -ExpectedContacts $test.ExpectedContacts `
                      -ExpectedFooters $test.ExpectedFooters
}

# Test line preservation (reconstruction)
Write-Host "Test Suite 2: Line Preservation" -ForegroundColor Cyan
Write-Host ""

foreach ($test in $testFiles) {
    Test-LinePreservation -TestName "$($test.Name) - Line Preservation" `
                          -FilePath $test.Path
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
    Write-Host "  Phase 1: ALL TESTS PASSED" -NoNewline
    Write-Host " " -NoNewline
    Write-Host "Pass" -ForegroundColor Green
}
else {
    Write-Host "  Phase 1: SOME TESTS FAILED" -ForegroundColor Red
}

Write-Host ""
