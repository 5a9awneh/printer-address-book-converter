<#
.SYNOPSIS
    Phase 3 Integration Tests - Output Writing & Pipeline

.DESCRIPTION
    Tests the refactored Phase 3 pipeline:
    - Write-AddressBook accepts normalized contacts
    - ConvertFrom-NormalizedContact maps to target brand fields
    - Full Convert-AddressBook pipeline works end-to-end
    - Validate-OutputFile checks output structure
#>

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Phase 3: Output Writing & Pipeline Integration Tests" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

$scriptPath = "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1"

# Test 1: Sharp → Canon (basic conversion)
Write-Host "[TEST 1] Sharp → Canon conversion" -ForegroundColor Yellow
$result1 = & $scriptPath -SourcePath "$PSScriptRoot\source_exports\10.52.18.61_SHARP_MX-3051.csv" -TargetBrand "Canon" -NoInteractive 2>&1
$outputFile1 = "$PSScriptRoot\source_exports\converted\10.52.18.61_SHARP_MX-3051_to_Canon.csv"

if (Test-Path $outputFile1) {
    $content1 = Get-Content $outputFile1
    if ($content1[0] -match "# Canon AddressBook CSV") {
        Write-Host "  [PASS] Canon header present" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Canon header missing" -ForegroundColor Red
    }
    
    $dataLines1 = $content1 | Where-Object { $_ -match "@" }
    if ($dataLines1.Count -ge 10) {
        Write-Host "  [PASS] Found $($dataLines1.Count) contact rows" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Only $($dataLines1.Count) contacts (expected >= 10)" -ForegroundColor Red
    }
} else {
    Write-Host "  [FAIL] Output file not created" -ForegroundColor Red
}

# Test 2: Sharp → Xerox (FirstName/LastName splitting)
Write-Host ""
Write-Host "[TEST 2] Sharp → Xerox (name splitting)" -ForegroundColor Yellow
$result2 = & $scriptPath -SourcePath "$PSScriptRoot\source_exports\10.52.18.101_SHARP_BP-50C31.csv" -TargetBrand "Xerox" -NoInteractive 2>&1
$outputFile2 = "$PSScriptRoot\source_exports\converted\10.52.18.101_SHARP_BP-50C31_to_Xerox.csv"

if (Test-Path $outputFile2) {
    $data2 = Import-Csv $outputFile2
    if ($data2.Count -ge 10) {
        Write-Host "  [PASS] Found $($data2.Count) contacts" -ForegroundColor Green
    }
    
    # Check FirstName/LastName columns exist
    if ($data2[0].PSObject.Properties['FirstName'] -and $data2[0].PSObject.Properties['LastName']) {
        Write-Host "  [PASS] FirstName/LastName columns present" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] FirstName/LastName columns missing" -ForegroundColor Red
    }
    
    # Check email field
    if ($data2[0].PSObject.Properties['E-mailAddress']) {
        Write-Host "  [PASS] E-mailAddress column present" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] E-mailAddress column missing" -ForegroundColor Red
    }
} else {
    Write-Host "  [FAIL] Output file not created" -ForegroundColor Red
}

# Test 3: Sharp → Develop (SearchKey generation)
Write-Host ""
Write-Host "[TEST 3] Sharp → Develop (SearchKey)" -ForegroundColor Yellow
$result3 = & $scriptPath -SourcePath "$PSScriptRoot\source_exports\10.52.30.242_SHARP_MX-5051.csv" -TargetBrand "Develop" -NoInteractive 2>&1
$outputFile3 = "$PSScriptRoot\source_exports\converted\10.52.30.242_SHARP_MX-5051_to_Develop.csv"

if (Test-Path $outputFile3) {
    $content3 = Get-Content $outputFile3 -Encoding Unicode
    if ($content3[0] -match "@Ver") {
        Write-Host "  [PASS] Develop header present" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Develop header missing" -ForegroundColor Red
    }
    
    $data3 = Import-Csv $outputFile3 -Delimiter "`t" -Encoding Unicode | Where-Object { $_.MailAddress }
    if ($data3.Count -ge 1) {
        Write-Host "  [PASS] Found $($data3.Count) contacts with email" -ForegroundColor Green
    }
    
    # Check SearchKey is populated
    if ($data3[0].SearchKey -and $data3[0].SearchKey -ne '') {
        Write-Host "  [PASS] SearchKey generated: $($data3[0].SearchKey)" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] SearchKey not generated" -ForegroundColor Red
    }
} else {
    Write-Host "  [FAIL] Output file not created" -ForegroundColor Red
}

# Test 4: Test Validate-OutputFile function
Write-Host ""
Write-Host "[TEST 4] Validate-OutputFile function" -ForegroundColor Yellow

# Load the Validate-OutputFile function
$scriptContent = Get-Content $scriptPath -Raw
$funcRegex = "(?ms)function Validate-OutputFile \{.*?^}"
$funcMatch = [regex]::Match($scriptContent, $funcRegex, [System.Text.RegularExpressions.RegexOptions]::Multiline)
if ($funcMatch.Success) {
    Invoke-Expression $funcMatch.Value
    
    # Also load dependencies
    $configRegex = '(?ms)\$Script:BrandConfig = @\{.*?^}'
    $configMatch = [regex]::Match($scriptContent, $configRegex)
    if ($configMatch.Success) {
        Invoke-Expression $configMatch.Value
    }
    
    $functions = @('Get-FileEncoding', 'Get-CsvStructure', 'Write-Log', 'Write-FunctionEntry', 'Write-FunctionExit')
    foreach ($funcName in $functions) {
        $funcRegex2 = "(?ms)function $funcName \{.*?^}"
        $funcMatch2 = [regex]::Match($scriptContent, $funcRegex2)
        if ($funcMatch2.Success) {
            Invoke-Expression $funcMatch2.Value
        }
    }
    
    $Script:LogFile = "test-phase3-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    
    # Validate a valid Canon output
    $validation1 = Validate-OutputFile -FilePath $outputFile1 -TargetBrand "Canon" -ExpectedContactCount 13
    if ($validation1.IsValid) {
        Write-Host "  [PASS] Canon output validated: $($validation1.ContactCount) contacts" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Canon output validation failed: $($validation1.Errors -join ', ')" -ForegroundColor Red
    }
    
    # Validate Xerox output
    $validation2 = Validate-OutputFile -FilePath $outputFile2 -TargetBrand "Xerox"
    if ($validation2.IsValid) {
        Write-Host "  [PASS] Xerox output validated: $($validation2.ContactCount) contacts" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Xerox output validation failed: $($validation2.Errors -join ', ')" -ForegroundColor Red
    }
    
    # Test validation on non-existent file
    $validation3 = Validate-OutputFile -FilePath "nonexistent.csv" -TargetBrand "Canon"
    if (-not $validation3.IsValid -and $validation3.Errors.Count -gt 0) {
        Write-Host "  [PASS] Correctly detected missing file" -ForegroundColor Green
    } else {
        Write-Host "  [FAIL] Did not detect missing file" -ForegroundColor Red
    }
    
} else {
    Write-Host "  [FAIL] Could not load Validate-OutputFile function" -ForegroundColor Red
}

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Phase 3 Tests Complete" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
