# Sanitize printer address book CSV files using a mock people list approach
# Preserves exact structure while replacing PII with synthetic data
#
# USAGE:
#   1. Edit the $sourceFiles hashtable below to point to your printer export files
#   2. Run: .\Sanitize-Samples.ps1
#   3. Sanitized files will be created in demo_samples/
#
# NOTE: This script is for creating demo samples from real exports.
#       End users don't need this - they browse/select their own files directly in the converter.

$ErrorActionPreference = "Stop"

# ============================================================================
# CONFIGURATION: Edit these paths to point to your own printer export files
# ============================================================================
$sourceFiles = @{
    Sharp   = "tests\source_samples\10.52.18.101_SHARP_BP-50C31.csv"
    Canon   = "tests\source_samples\10.52.18.11_Canon_iR-ADV C3930.csv"
    Xerox   = "tests\source_samples\10.52.18.148_Xerox VersaLink C7130.csv"
    Develop = "tests\source_samples\10.52.18.40_Develop ineo+ 224e.csv"
}

# Outlook test file for testing "Create from Outlook list" mode
$outlookTestFile = "$PSScriptRoot\demo_samples\Outlook_Sample.txt"

# Verify source files exist
foreach ($brand in $sourceFiles.Keys) {
    if (-not (Test-Path $sourceFiles[$brand])) {
        Write-Host "ERROR: Source file not found for $brand" -ForegroundColor Red
        Write-Host "  Path: $($sourceFiles[$brand])" -ForegroundColor Yellow
        Write-Host "  Please edit the `$sourceFiles hashtable in this script" -ForegroundColor Yellow
        exit 1
    }
}

# ============================================================================
# Mock people list - consistent across all samples
# ============================================================================
$mockPeople = @(
    @{ Name = "John Smith"; Email = "john.smith@example.com" }
    @{ Name = "Jane Doe"; Email = "jane.doe@example.com" }
    @{ Name = "Michael Johnson"; Email = "michael.johnson@example.com" }
    @{ Name = "Emily Davis"; Email = "emily.davis@example.com" }
    @{ Name = "Robert Wilson"; Email = "robert.wilson@example.com" }
    @{ Name = "Sarah Brown"; Email = "sarah.brown@example.com" }
    @{ Name = "David Miller"; Email = "david.miller@example.com" }
    @{ Name = "Lisa Anderson"; Email = "lisa.anderson@example.com" }
    @{ Name = "Mark Taylor"; Email = "mark.taylor@example.com" }
    @{ Name = "Susan Clark"; Email = "susan.clark@example.com" }
    @{ Name = "Tom Harris"; Email = "tom.harris@example.com" }
    @{ Name = "Nancy White"; Email = "nancy.white@example.com" }
    @{ Name = "Chris Martin"; Email = "chris.martin@example.com" }
    @{ Name = "Patricia Garcia"; Email = "patricia.garcia@example.com" }
    @{ Name = "James Rodriguez"; Email = "james.rodriguez@example.com" }
    @{ Name = "Jennifer Martinez"; Email = "jennifer.martinez@example.com" }
    @{ Name = "William Lee"; Email = "william.lee@example.com" }
    @{ Name = "Linda Walker"; Email = "linda.walker@example.com" }
    @{ Name = "Richard Hall"; Email = "richard.hall@example.com" }
    @{ Name = "Charles Young"; Email = "charles.young@example.com" }
    @{ Name = "Barbara Edwards"; Email = "barbara.edwards@example.com" }
    @{ Name = "George Adams"; Email = "george.adams@example.com" }
    @{ Name = "Dorothy Foster"; Email = "dorothy.foster@example.com" }
    @{ Name = "Henry Greene"; Email = "henry.greene@example.com" }
    @{ Name = "Alice Harris"; Email = "alice.harris@example.com" }
)

Write-Host "Starting CSV sanitization using email detection..." -ForegroundColor Cyan
Write-Host ""

# Email pattern - same as main converter uses
$emailPattern = '[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'

# SearchKey generator for Develop format - same as main converter
function Get-SearchKey {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return 'Other'
    }

    $first = $Name.Substring(0, 1).ToUpper()

    switch -Regex ($first) {
        '[A-C]' { return 'Abc' }
        '[D-F]' { return 'Def' }
        '[G-I]' { return 'Ghi' }
        '[J-L]' { return 'Jkl' }
        '[M-O]' { return 'Mno' }
        '[P-S]' { return 'Pqrs' }
        '[T-V]' { return 'Tuv' }
        '[W-Z]' { return 'Wxyz' }
        default { return 'Other' }
    }
}

# Universal sanitization function - works for ALL brands
function Sanitize-File {
    param(
        [string]$SourcePath,
        [string]$DestPath,
        [string]$BrandName
    )
    
    Write-Host "Processing $BrandName..." -NoNewline
    
    # Detect encoding
    $bytes = Get-Content $SourcePath -Encoding Byte -TotalCount 4
    if ($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        $encodingName = 'Unicode'
        $encoding = [System.Text.Encoding]::Unicode
        $delimiter = "`t"
    } else {
        $encodingName = 'UTF8'
        $encoding = [System.Text.Encoding]::UTF8
        $delimiter = ','
    }
    
    $allLines = Get-Content $SourcePath -Encoding $encodingName
    $outputLines = @()
    $mockIndex = 0
    
    foreach ($line in $allLines) {
        # Check if line contains email (= data row)
        if ($line -match $emailPattern) {
            if ($mockIndex -lt $mockPeople.Count) {
                $mockPerson = $mockPeople[$mockIndex]
                
                # Replace the email
                $sanitized = $line -replace $emailPattern, $mockPerson.Email
                
                # Replace ALL name-like content:
                # For CSV (comma-delimited): Replace all quoted text that isn't a special keyword/value
                if ($delimiter -eq ',') {
                    # Replace quoted strings that look like names (contain letters, not just numbers/symbols)
                    # Exclude common keywords
                    $keywords = @('TRUE', 'FALSE', 'Yes', 'No', 'email', 'data', 'smtp', 'off', 'encrypted2')
                    $keywordPattern = ($keywords | ForEach-Object { [regex]::Escape($_) }) -join '|'
                    
                    # Find and replace all quoted fields
                    $sanitized = [regex]::Replace($sanitized, '"([^"]*)"', {
                        param($match)
                        $value = $match.Groups[1].Value
                        # Keep if it's:
                        # - A keyword (TRUE, FALSE, etc.)
                        # - Only numbers
                        # - Empty string
                        # - Base64-like (contains +/= and is longer than typical name)
                        # - Contains @ (email, already replaced) or path separators
                        # - Contains "encodingMethod" in context
                        if ($value -match "^($keywordPattern)`$|^\d+`$|^`$|^[A-Za-z0-9+/=]{20,}`$|[@/\\]|encodingMethod") {
                            return $match.Value
                        }
                        # Replace if it contains letters (likely a name)
                        if ($value -match '[A-Za-z]') {
                            return '"' + $mockPerson.Name + '"'
                        }
                        return $match.Value
                    })
                }
                # For TSV (tab-delimited): Replace specific field positions known to be names
                else {
                    $fields = $sanitized -split "`t"
                    if ($fields.Count -gt 11) {
                        # Develop format: field 1=Name, field 3=Furigana, field 4=SearchKey
                        $fields[1] = $mockPerson.Name
                        $fields[3] = $mockPerson.Name
                        $fields[4] = Get-SearchKey -Name $mockPerson.Name
                    }
                    $sanitized = $fields -join "`t"
                }
                
                $outputLines += $sanitized
                $mockIndex++
            }
        }
        else {
            # Header/footer line - sanitize domain references only
            $sanitized = $line -replace '@[a-zA-Z0-9.-]+\.(org|com)', '@example.com'
            $sanitized = $sanitized -replace '\b[A-Z]{3,}\b', 'Example Org'
            $outputLines += $sanitized
        }
    }
    
    $content = $outputLines -join "`r`n"
    [System.IO.File]::WriteAllText($DestPath, $content, $encoding)
    Write-Host " Done ($mockIndex rows)" -ForegroundColor Green
}

# Process all source files
Sanitize-File -SourcePath $sourceFiles.Sharp -DestPath "$PSScriptRoot\demo_samples\Sharp_Sample.csv" -BrandName "Sharp"
Sanitize-File -SourcePath $sourceFiles.Canon -DestPath "$PSScriptRoot\demo_samples\Canon_Sample.csv" -BrandName "Canon"
Sanitize-File -SourcePath $sourceFiles.Xerox -DestPath "$PSScriptRoot\demo_samples\Xerox_Sample.csv" -BrandName "Xerox"
Sanitize-File -SourcePath $sourceFiles.Develop -DestPath "$PSScriptRoot\demo_samples\Develop_Sample.csv" -BrandName "Develop"

Write-Host ""
Write-Host "Verifying sanitization..." -ForegroundColor Cyan
Write-Host ""

$samples = @(
    @{ Name = "Sharp"; Path = "$PSScriptRoot\demo_samples\Sharp_Sample.csv" }
    @{ Name = "Canon"; Path = "$PSScriptRoot\demo_samples\Canon_Sample.csv" }
    @{ Name = "Xerox"; Path = "$PSScriptRoot\demo_samples\Xerox_Sample.csv" }
    @{ Name = "Develop"; Path = "$PSScriptRoot\demo_samples\Develop_Sample.csv" }
)

$allPassed = $true

foreach ($sample in $samples) {
    Write-Host "Checking $($sample.Name)..." -NoNewline
    $content = Get-Content $sample.Path -Raw
    
    # Check for PII patterns - verify only mock data remains
    $hasPII = $false
    # Check if any email NOT from example.com domain exists
    if ($content -match '[a-zA-Z0-9._%+-]+@(?!example\.com)[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}') {
        $hasPII = $true
    }
    # Verify only mock names are present
    $mockNames = ($mockPeople | ForEach-Object { [regex]::Escape($_.Name) }) -join '|'
    $foundNames = [regex]::Matches($content, '[A-Z][a-z]+ [A-Z][a-z]+')
    foreach ($name in $foundNames) {
        if ($name.Value -notmatch "^($mockNames)`$") {
            $hasPII = $true
            break
        }
    }
    
    if ($hasPII) {
        Write-Host " FAIL - PII detected" -ForegroundColor Red
        $allPassed = $false
    }
    else {
        Write-Host " PASS" -ForegroundColor Green
    }
}

Write-Host ""
if ($allPassed) {
    Write-Host "All sanitization checks passed!" -ForegroundColor Green
} else {
    Write-Host "Some sanitization checks failed!" -ForegroundColor Red
    exit 1
}

Write-Host ""
Write-Host "Running conversion tests..." -ForegroundColor Yellow
Write-Host ""

$tests = @(
    @{ Source = "Sharp_Sample.csv"; Target = "Canon"; Description = "Sharp -> Canon" }
    @{ Source = "Canon_Sample.csv"; Target = "Sharp"; Description = "Canon -> Sharp" }
    @{ Source = "Xerox_Sample.csv"; Target = "Develop"; Description = "Xerox -> Develop" }
    @{ Source = "Develop_Sample.csv"; Target = "Canon"; Description = "Develop -> Canon" }
)

foreach ($test in $tests) {
    Write-Host "Test: $($test.Description)..." -NoNewline
    
    $result = & "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1" `
        -SourcePath "$PSScriptRoot\demo_samples\$($test.Source)" `
        -TargetBrand $test.Target `
        -NoInteractive 2>&1 | Out-String
    
    if ($result -match "Converted: \d+") {
        Write-Host " PASS" -ForegroundColor Green
    } else {
        Write-Host " FAIL" -ForegroundColor Red
        Write-Host "Output: $result"
    }
}

Write-Host ""
Write-Host "Sanitization and testing complete!" -ForegroundColor Green
