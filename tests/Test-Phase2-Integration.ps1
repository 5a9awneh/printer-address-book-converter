<#
.SYNOPSIS
    Phase 2.3 Integration Tests - Multi-Brand Conversions

.DESCRIPTION
    Tests end-to-end conversion pipeline using real CSV files:
    - Read from source brand → Normalize → Convert to target brand
    - Verify data integrity across all brand-pair combinations
    - Validate email and name field mapping accuracy
#>

param(
    [switch]$Verbose
)

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Phase 2.3: Multi-Brand Integration Tests" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

$testFiles = @(
    @{ File = '10.52.18.11_Canon_iR-ADV C3930.csv'; Brand = 'Canon'; ExpectedContacts = 17 }
    @{ File = '10.52.18.101_SHARP_BP-50C31.csv'; Brand = 'Sharp'; ExpectedContacts = 15 }
    @{ File = '10.52.18.148_Xerox VersaLink C7130.csv'; Brand = 'Xerox'; ExpectedContacts = 4 }
    @{ File = '10.52.18.144_bizhub C360i.csv'; Brand = 'Develop'; ExpectedContacts = 16 }
)

$targetBrands = @('Canon', 'Sharp', 'Xerox', 'Develop')

# Load main script to get functions
$scriptPath = "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1"
$scriptContent = Get-Content $scriptPath -Raw

# Load configuration
$configRegex = '(?ms)\$Script:BrandConfig = @\{.*?^}'
$configMatch = [regex]::Match($scriptContent, $configRegex, [System.Text.RegularExpressions.RegexOptions]::Multiline)
if ($configMatch.Success) {
    Invoke-Expression $configMatch.Value
}

# Load required functions
$functions = @(
    'Write-Log', 'Write-FunctionEntry', 'Write-FunctionExit',
    'Test-Email', 'Split-FullName', 'Get-FileEncoding', 'Get-CsvStructure',
    'ConvertTo-NormalizedContact', 'ConvertFrom-NormalizedContact'
)

$Script:LogFile = "test-phase2-3-$(Get-Date -Format 'yyyy-MM-dd-HHmmss').log"

foreach ($funcName in $functions) {
    $funcRegex = "(?ms)function $funcName \{.*?^}"
    $funcMatch = [regex]::Match($scriptContent, $funcRegex, [System.Text.RegularExpressions.RegexOptions]::Multiline)
    if ($funcMatch.Success) {
        Invoke-Expression $funcMatch.Value
    }
}

$totalTests = 0
$passedTests = 0

foreach ($testFile in $testFiles) {
    $sourcePath = Join-Path "$PSScriptRoot\source_exports" $testFile.File
    
    if (-not (Test-Path $sourcePath)) {
        Write-Host "⚠ Skipping $($testFile.File) - file not found" -ForegroundColor Yellow
        continue
    }
    
    $sourceBrand = $testFile.Brand
    
    Write-Host "Source: $sourceBrand - $($testFile.File)" -ForegroundColor Cyan
    
    # Read and parse source file
    try {
        $encoding = Get-FileEncoding -FilePath $sourcePath
        $structure = Get-CsvStructure -FilePath $sourcePath -Encoding $encoding
        
        # Parse contacts based on brand
        $config = $Script:BrandConfig[$sourceBrand]
        
        if ($sourceBrand -eq 'Develop') {
            # Extract header (line 3: AbbrNo, Name, ...)
            $headerLine = $structure.Headers | Where-Object { $_ -match '^AbbrNo' } | Select-Object -First 1
            
            # Filter out metadata (@Ver, @End) and "alternative" definition row
            $dataLines = $structure.Contacts | Where-Object { 
                $_ -notmatch '^@(Ver|End)' -and 
                $_ -notmatch '^"alternative"' -and 
                -not [string]::IsNullOrWhiteSpace($_)
            }
            
            # Combine header + data
            $allLines = @($headerLine) + $dataLines
            
            $tempFile = [System.IO.Path]::GetTempFileName()
            $allLines | Out-File -FilePath $tempFile -Encoding Unicode
            $csvData = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        elseif ($sourceBrand -eq 'Canon') {
            # Extract header line (last non-comment line before contacts)
            $headerLine = $structure.Headers | Where-Object { $_ -notmatch '^\s*#' -and -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1
            $csvLines = $structure.Contacts | Where-Object { -not ($_ -match '^\s*#') -and -not ([string]::IsNullOrWhiteSpace($_)) }
            
            # Combine header + contacts
            $allLines = @($headerLine) + $csvLines
            
            $tempFile = [System.IO.Path]::GetTempFileName()
            $allLines | Out-File -FilePath $tempFile -Encoding UTF8
            $csvData = Import-Csv -Path $tempFile -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        else {
            # Xerox and Sharp: Header is last line of Headers section (before contacts with @)
            $headerLine = $structure.Headers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Last 1
            
            # Combine header + contacts
            $allLines = @($headerLine) + $structure.Contacts
            
            $tempFile = [System.IO.Path]::GetTempFileName()
            $allLines | Out-File -FilePath $tempFile -Encoding $encoding
            $csvData = Import-Csv -Path $tempFile -Encoding $encoding -Delimiter $config.Delimiter
            Remove-Item -Path $tempFile -Force
        }
        
        # Normalize all contacts
        $normalized = @()
        foreach ($row in $csvData) {
            $norm = ConvertTo-NormalizedContact -Contact $row -SourceBrand $sourceBrand
            if ($norm) {
                $normalized += $norm
            }
        }
        
        Write-Host "  Parsed: $($normalized.Count) contacts (expected: $($testFile.ExpectedContacts))" -ForegroundColor Gray
        
        # Test conversion to each target brand
        foreach ($targetBrand in $targetBrands) {
            if ($targetBrand -eq $sourceBrand) {
                continue  # Skip same-brand conversion
            }
            
            $testName = "$sourceBrand → $targetBrand"
            $totalTests++
            
            $converted = @()
            $failed = 0
            
            foreach ($norm in $normalized) {
                $target = ConvertFrom-NormalizedContact -NormalizedContact $norm -TargetBrand $targetBrand
                if ($target) {
                    $converted += $target
                }
                else {
                    $failed++
                }
            }
            
            # Verify all contacts converted successfully
            if ($converted.Count -eq $normalized.Count -and $failed -eq 0) {
                # Verify target format has correct fields
                $targetConfig = $Script:BrandConfig[$targetBrand]
                $firstContact = $converted[0]
                
                $hasEmail = $firstContact.ContainsKey($targetConfig.OutputFields.Email)
                $hasName = $firstContact.ContainsKey($targetConfig.OutputFields.DisplayName)
                
                if ($hasEmail -and $hasName) {
                    Write-Host "  [PASS] $testName : $($converted.Count) contacts" -ForegroundColor Green
                    $passedTests++
                }
                else {
                    Write-Host "  [FAIL] $testName : missing fields" -ForegroundColor Red
                    if ($Verbose) {
                        Write-Host "    hasEmail=$hasEmail, hasName=$hasName" -ForegroundColor Yellow
                    }
                }
            }
            else {
                Write-Host "  [FAIL] $testName : $($converted.Count)/$($normalized.Count) converted, $failed failed" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "  [ERROR] Processing file: $_" -ForegroundColor Red
    }
    
    Write-Host ""
}

# Summary
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Test Summary" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total Brand-Pair Tests: $totalTests"
Write-Host "  Passed: $passedTests" -ForegroundColor Green
Write-Host "  Failed: $($totalTests - $passedTests)" -ForegroundColor $(if ($passedTests -eq $totalTests) { 'Green' } else { 'Red' })
Write-Host ""

if ($passedTests -eq $totalTests) {
    Write-Host "  Phase 2.3: ALL TESTS PASSED" -ForegroundColor Green
    Write-Host "  All brand-pair combinations validated!" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "  Phase 2.3: SOME TESTS FAILED" -ForegroundColor Red
    exit 1
}
