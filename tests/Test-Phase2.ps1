<#
.SYNOPSIS
    Phase 2 Test Suite - Field Mapping Engine

.DESCRIPTION
    Tests for ConvertTo-NormalizedContact() function to verify:
    - Contact normalization from all brand formats
    - Email and name field extraction
    - FirstName/LastName parsing
    - Validation (email format, missing fields)
#>

Write-Host ""
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Phase 2: Field Mapping Tests" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""

# Load script functions without running Main
$scriptPath = "$PSScriptRoot\..\Convert-PrinterAddressBook.ps1"
$scriptContent = Get-Content $scriptPath -Raw

# Extract function definitions only (not Main execution)
$functions = @(
    'Write-Log',
    'Write-FunctionEntry',
    'Write-FunctionExit',
    'Test-Email',
    'Split-FullName',
    'ConvertTo-NormalizedContact',
    'ConvertFrom-NormalizedContact'
)

# Load configuration
$configRegex = '(?ms)\$Script:BrandConfig = @\{.*?\n\}'
$config = [regex]::Match($scriptContent, $configRegex).Value
if ($config) {
    Invoke-Expression $config
}

# Load log file variable
$Script:LogFile = "test-phase2-$(Get-Date -Format 'yyyy-MM-dd').log"

# Load functions
foreach ($funcName in $functions) {
    $funcRegex = "(?ms)function $funcName \{.*?^}"
    $funcMatch = [regex]::Match($scriptContent, $funcRegex, [System.Text.RegularExpressions.RegexOptions]::Multiline)
    if ($funcMatch.Success) {
        Invoke-Expression $funcMatch.Value
    }
}

$totalTests = 0
$passedTests = 0

# Test 1: Canon contact normalization
Write-Host "Test 1: Canon Contact Normalization" -ForegroundColor Cyan

$canonContact = [PSCustomObject]@{
    cn = 'John Smith'
    mailaddress = 'john.smith@example.com'
    objectclass = 'email'
}

$result = ConvertTo-NormalizedContact -Contact $canonContact -SourceBrand 'Canon'
$totalTests++

if ($result -and 
    $result.Email -eq 'john.smith@example.com' -and
    $result.DisplayName -eq 'John Smith' -and
    $result.FirstName -eq 'John' -and
    $result.LastName -eq 'Smith') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
    Write-Host "  Expected: John Smith <john.smith@example.com>" -ForegroundColor Yellow
    if ($result) {
        Write-Host "  Got: $($result.DisplayName) <$($result.Email)>" -ForegroundColor Yellow
    }
    else {
        Write-Host "  Got: null" -ForegroundColor Yellow
    }
}
Write-Host ""

# Test 2: Sharp contact normalization
Write-Host "Test 2: Sharp Contact Normalization" -ForegroundColor Cyan

$sharpContact = [PSCustomObject]@{
    name = 'Jane Doe'
    'mail-address' = 'jane.doe@example.com'
}

$result = ConvertTo-NormalizedContact -Contact $sharpContact -SourceBrand 'Sharp'
$totalTests++

if ($result -and 
    $result.Email -eq 'jane.doe@example.com' -and
    $result.DisplayName -eq 'Jane Doe' -and
    $result.FirstName -eq 'Jane' -and
    $result.LastName -eq 'Doe') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Test 3: Xerox contact with DisplayName
Write-Host "Test 3: Xerox Contact with DisplayName" -ForegroundColor Cyan

$xeroxContact = [PSCustomObject]@{
    DisplayName = 'Robert Johnson'
    'E-mailAddress' = 'robert.j@example.com'
    FirstName = ''
    LastName = ''
}

$result = ConvertTo-NormalizedContact -Contact $xeroxContact -SourceBrand 'Xerox'
$totalTests++

if ($result -and 
    $result.Email -eq 'robert.j@example.com' -and
    $result.DisplayName -eq 'Robert Johnson' -and
    $result.FirstName -eq 'Robert' -and
    $result.LastName -eq 'Johnson') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Test 4: Xerox contact with FirstName/LastName
Write-Host "Test 4: Xerox Contact with FirstName/LastName" -ForegroundColor Cyan

$xeroxContact2 = [PSCustomObject]@{
    DisplayName = ''
    'E-mailAddress' = 'mary.williams@example.com'
    FirstName = 'Mary'
    LastName = 'Williams'
}

$result = ConvertTo-NormalizedContact -Contact $xeroxContact2 -SourceBrand 'Xerox'
$totalTests++

if ($result -and 
    $result.Email -eq 'mary.williams@example.com' -and
    $result.DisplayName -eq 'Mary Williams' -and
    $result.FirstName -eq 'Mary' -and
    $result.LastName -eq 'Williams') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Test 5: Develop contact
Write-Host "Test 5: Develop Contact Normalization" -ForegroundColor Cyan

$developContact = [PSCustomObject]@{
    Name = 'David Brown'
    MailAddress = 'david.brown@example.com'
}

$result = ConvertTo-NormalizedContact -Contact $developContact -SourceBrand 'Develop'
$totalTests++

if ($result -and 
    $result.Email -eq 'david.brown@example.com' -and
    $result.DisplayName -eq 'David Brown' -and
    $result.FirstName -eq 'David' -and
    $result.LastName -eq 'Brown') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Test 6: Invalid email rejection
Write-Host "Test 6: Invalid Email Rejection" -ForegroundColor Cyan

$invalidContact = [PSCustomObject]@{
    cn = 'Invalid User'
    mailaddress = 'not-an-email'
}

$result = ConvertTo-NormalizedContact -Contact $invalidContact -SourceBrand 'Canon'
$totalTests++

if ($null -eq $result) {
    Write-Host "  Result: PASS (correctly rejected)" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL (should have rejected)" -ForegroundColor Red
}
Write-Host ""

# Test 7: Missing name rejection
Write-Host "Test 7: Missing Name Rejection" -ForegroundColor Cyan

$noNameContact = [PSCustomObject]@{
    cn = ''
    mailaddress = 'test@example.com'
}

$result = ConvertTo-NormalizedContact -Contact $noNameContact -SourceBrand 'Canon'
$totalTests++

if ($null -eq $result) {
    Write-Host "  Result: PASS (correctly rejected)" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL (should have rejected)" -ForegroundColor Red
}
Write-Host ""

# Test 8: Comma-separated name parsing (LastName, FirstName)
Write-Host "Test 8: Comma-Separated Name Parsing" -ForegroundColor Cyan

$commaNameContact = [PSCustomObject]@{
    cn = 'Smith, Alice'
    mailaddress = 'alice.smith@example.com'
}

$result = ConvertTo-NormalizedContact -Contact $commaNameContact -SourceBrand 'Canon'
$totalTests++

if ($result -and 
    $result.FirstName -eq 'Alice' -and
    $result.LastName -eq 'Smith' -and
    $result.DisplayName -eq 'Smith, Alice') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
    if ($result) {
        Write-Host "  Expected: FirstName=Alice, LastName=Smith" -ForegroundColor Yellow
        Write-Host "  Got: FirstName=$($result.FirstName), LastName=$($result.LastName)" -ForegroundColor Yellow
    }
}
Write-Host ""

# Test 9: Canon to Sharp field mapping
Write-Host "Test 9: Canon to Sharp Field Mapping" -ForegroundColor Cyan

$canonContact = [PSCustomObject]@{
    cn = 'Test User'
    mailaddress = 'test@example.com'
}

$normalized = ConvertTo-NormalizedContact -Contact $canonContact -SourceBrand 'Canon'
$sharpFormat = ConvertFrom-NormalizedContact -NormalizedContact $normalized -TargetBrand 'Sharp'
$totalTests++

if ($sharpFormat -and 
    $sharpFormat['name'] -eq 'Test User' -and
    $sharpFormat['mail-address'] -eq 'test@example.com') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
    if ($sharpFormat) {
        Write-Host "  Got: name=$($sharpFormat['name']), mail-address=$($sharpFormat['mail-address'])" -ForegroundColor Yellow
    }
}
Write-Host ""

# Test 10: Sharp to Xerox field mapping
Write-Host "Test 10: Sharp to Xerox Field Mapping" -ForegroundColor Cyan

$sharpContact = [PSCustomObject]@{
    name = 'Jane Doe'
    'mail-address' = 'jane@example.com'
}

$normalized = ConvertTo-NormalizedContact -Contact $sharpContact -SourceBrand 'Sharp'
$xeroxFormat = ConvertFrom-NormalizedContact -NormalizedContact $normalized -TargetBrand 'Xerox'
$totalTests++

if ($xeroxFormat -and 
    $xeroxFormat['DisplayName'] -eq 'Jane Doe' -and
    $xeroxFormat['FirstName'] -eq 'Jane' -and
    $xeroxFormat['LastName'] -eq 'Doe' -and
    $xeroxFormat['E-mailAddress'] -eq 'jane@example.com') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Test 11: Xerox to Develop field mapping
Write-Host "Test 11: Xerox to Develop Field Mapping" -ForegroundColor Cyan

$xeroxContact = [PSCustomObject]@{
    DisplayName = 'Bob Smith'
    'E-mailAddress' = 'bob@example.com'
    FirstName = 'Bob'
    LastName = 'Smith'
}

$normalized = ConvertTo-NormalizedContact -Contact $xeroxContact -SourceBrand 'Xerox'
$developFormat = ConvertFrom-NormalizedContact -NormalizedContact $normalized -TargetBrand 'Develop'
$totalTests++

if ($developFormat -and 
    $developFormat['Name'] -eq 'Bob Smith' -and
    $developFormat['MailAddress'] -eq 'bob@example.com') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Test 12: Develop to Canon field mapping
Write-Host "Test 12: Develop to Canon Field Mapping" -ForegroundColor Cyan

$developContact = [PSCustomObject]@{
    Name = 'Alice Johnson'
    MailAddress = 'alice@example.com'
}

$normalized = ConvertTo-NormalizedContact -Contact $developContact -SourceBrand 'Develop'
$canonFormat = ConvertFrom-NormalizedContact -NormalizedContact $normalized -TargetBrand 'Canon'
$totalTests++

if ($canonFormat -and 
    $canonFormat['cn'] -eq 'Alice Johnson' -and
    $canonFormat['mailaddress'] -eq 'alice@example.com') {
    Write-Host "  Result: PASS" -ForegroundColor Green
    $passedTests++
}
else {
    Write-Host "  Result: FAIL" -ForegroundColor Red
}
Write-Host ""

# Summary
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host "  Test Summary" -ForegroundColor Cyan
Write-Host "===============================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total Tests: $totalTests"
Write-Host "  Passed: $passedTests" -ForegroundColor Green
Write-Host "  Failed: $($totalTests - $passedTests)" -ForegroundColor $(if ($passedTests -eq $totalTests) { 'Green' } else { 'Red' })
Write-Host ""

if ($passedTests -eq $totalTests) {
    Write-Host "  Phase 2 (2.1-2.2): ALL TESTS PASSED" -ForegroundColor Green
    exit 0
}
else {
    Write-Host "  Phase 2 (2.1-2.2): SOME TESTS FAILED" -ForegroundColor Red
    exit 1
}
