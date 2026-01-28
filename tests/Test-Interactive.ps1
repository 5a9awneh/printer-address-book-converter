<#
.SYNOPSIS
    Interactive Workflow Testing - Manual UI/UX Validation

.DESCRIPTION
    Guided manual test for validating interactive user workflows:
    - File selection dialogs (foreground behavior)
    - Menu navigation and mode selection
    - Single/Batch/Merge conversion workflows
    - Backup and restore validation
    - Success/error messaging
    
    This test requires manual interaction and provides step-by-step
    instructions with validation checkpoints.

.NOTES
    Task 2.4: Moved from Phase 4.4 due to dialog display issues
    Manual execution required - validates actual user experience
#>

param(
    [switch]$Quick  # Skip detailed validation steps
)

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Interactive Workflow Testing" -ForegroundColor Cyan
Write-Host "  Manual UI/UX Validation for Printer Address Book Converter" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "This test requires manual interaction to validate the user experience." -ForegroundColor Yellow
Write-Host "Follow the instructions and confirm each step works correctly." -ForegroundColor Yellow
Write-Host ""

$testResults = @{
    TotalTests = 0
    Passed = 0
    Failed = 0
    Skipped = 0
}

function Test-Step {
    param(
        [string]$TestName,
        [string]$Instructions,
        [string]$ExpectedResult,
        [switch]$AutoPass
    )
    
    $script:testResults.TotalTests++
    
    Write-Host ""
    Write-Host "─────────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "Test: $TestName" -ForegroundColor Cyan
    Write-Host "─────────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host ""
    Write-Host "Instructions:" -ForegroundColor Yellow
    Write-Host "  $Instructions"
    Write-Host ""
    Write-Host "Expected Result:" -ForegroundColor Yellow
    Write-Host "  $ExpectedResult"
    Write-Host ""
    
    if ($AutoPass) {
        Write-Host "[AUTO-PASS] Test skipped in quick mode" -ForegroundColor Gray
        $script:testResults.Skipped++
        return
    }
    
    $response = Read-Host "Did this test PASS? (Y/N/S=Skip)"
    
    switch ($response.ToUpper()) {
        'Y' {
            Write-Host "✓ PASS" -ForegroundColor Green
            $script:testResults.Passed++
        }
        'N' {
            Write-Host "✗ FAIL" -ForegroundColor Red
            $details = Read-Host "Describe the issue (optional)"
            if ($details) {
                Write-Host "  Issue: $details" -ForegroundColor Yellow
            }
            $script:testResults.Failed++
        }
        'S' {
            Write-Host "⊘ SKIPPED" -ForegroundColor Gray
            $script:testResults.Skipped++
        }
        default {
            Write-Host "⊘ SKIPPED (invalid response)" -ForegroundColor Gray
            $script:testResults.Skipped++
        }
    }
}

# Test 1: Script Launch
Test-Step `
    -TestName "Script Launch" `
    -Instructions "Run: .\Convert-PrinterAddressBook.ps1 (without parameters)" `
    -ExpectedResult "Welcome banner appears with version number. Main menu displays with 4 options: Single/Batch/Merge/Exit. Cursor is ready for input." `
    -AutoPass:$Quick

# Test 2: File Dialog - Single File
Test-Step `
    -TestName "File Selection Dialog - Single File" `
    -Instructions "From main menu, select 'Single File Conversion'. File selection dialog should appear." `
    -ExpectedResult "Dialog appears IN FOREGROUND (no alt+tab needed). Can navigate to tests/source_exports/. File list is visible and selectable. Dialog is responsive." `
    -AutoPass:$Quick

# Test 3: Single File Conversion
Test-Step `
    -TestName "Single File Conversion Workflow" `
    -Instructions "Select any Canon CSV file (e.g., 10.52.18.11_Canon_iR-ADV C3930.csv). Choose Sharp as target brand. Let conversion complete." `
    -ExpectedResult "Source brand auto-detected (Canon). Target brand menu appears. Conversion executes with progress messages. Success message shows output location. File is created in 'converted' folder." `
    -AutoPass:$Quick

# Test 4: Batch File Selection
Test-Step `
    -TestName "File Selection Dialog - Multiple Files" `
    -Instructions "From main menu, select 'Batch Convert Multiple Files'. Multi-select dialog should appear." `
    -ExpectedResult "Dialog allows selecting MULTIPLE files (Ctrl+Click). Dialog appears in foreground. Selected files are highlighted. OK button accepts multiple selections." `
    -AutoPass:$Quick

# Test 5: Batch Conversion
Test-Step `
    -TestName "Batch Conversion Workflow" `
    -Instructions "Select 2-3 different brand files (e.g., Canon, Sharp, Xerox). Choose a target brand. Let batch conversion complete." `
    -ExpectedResult "Each file processed sequentially. Progress shown for each file. Source brands auto-detected correctly. All conversions succeed. Summary shows count of successful conversions." `
    -AutoPass:$Quick

# Test 6: Merge Mode
Test-Step `
    -TestName "Merge Multiple Files Workflow" `
    -Instructions "From main menu, select 'Merge Multiple Files'. Select 2-3 Canon files. Choose Canon as target." `
    -ExpectedResult "Multi-select dialog appears. Files are merged into single output. Duplicate contacts are handled (or warned). Output contains all contacts from all sources." `
    -AutoPass:$Quick

# Test 7: Backup Verification
Test-Step `
    -TestName "Backup File Creation" `
    -Instructions "Check tests/source_exports/backup/ folder after previous conversions." `
    -ExpectedResult "Backup files exist with timestamps (YYYY-MM-DD-HHMMSS format). Original files are unchanged. Backup files are valid CSVs matching original content." `
    -AutoPass:$Quick

# Test 8: Error Handling - Cancel Dialog
Test-Step `
    -TestName "Cancel File Selection" `
    -Instructions "Launch script, select Single mode, then CANCEL the file dialog." `
    -ExpectedResult "Script displays 'No file selected' message (in Yellow). Returns to main menu gracefully (no crash/error). Can proceed with another operation." `
    -AutoPass:$Quick

# Test 9: Error Handling - Invalid File
Test-Step `
    -TestName "Invalid File Handling" `
    -Instructions "Use non-interactive mode: .\Convert-PrinterAddressBook.ps1 -SourceFile 'nonexistent.csv' -TargetBrand Canon" `
    -ExpectedResult "Error message displayed (file not found). Error is descriptive and helpful. Script exits cleanly with error code." `
    -AutoPass:$Quick

# Test 10: Help Text
Test-Step `
    -TestName "Help Documentation" `
    -Instructions "Run: Get-Help .\Convert-PrinterAddressBook.ps1 -Detailed" `
    -ExpectedResult "Synopsis, description, parameter details are shown. Examples include both interactive and non-interactive usage. Text is clear and helpful." `
    -AutoPass:$Quick

# Test 11: Exit Menu
Test-Step `
    -TestName "Exit Menu Option" `
    -Instructions "Launch script, press ESC or select 'Exit' option." `
    -ExpectedResult "Script displays 'Cancelled' message. Exits cleanly. No errors in console. Returns to PowerShell prompt." `
    -AutoPass:$Quick

# Test 12: Brand Auto-Detection
Test-Step `
    -TestName "Brand Detection Accuracy" `
    -Instructions "Test with one file from each brand (Canon, Sharp, Xerox, Develop) in Single mode." `
    -ExpectedResult "Each brand is correctly detected. Detection message shows before conversion. No misidentification occurs." `
    -AutoPass:$Quick

# Summary
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Test Summary" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Total Tests: $($testResults.TotalTests)"
Write-Host "  Passed: $($testResults.Passed)" -ForegroundColor Green
Write-Host "  Failed: $($testResults.Failed)" -ForegroundColor $(if ($testResults.Failed -eq 0) { 'Green' } else { 'Red' })
Write-Host "  Skipped: $($testResults.Skipped)" -ForegroundColor Gray
Write-Host ""

if ($testResults.Failed -eq 0 -and $testResults.Passed -gt 0) {
    Write-Host "  Interactive Testing: ALL PASSED" -ForegroundColor Green
    Write-Host "  UI/UX validation complete!" -ForegroundColor Green
    exit 0
}
elseif ($testResults.Failed -gt 0) {
    Write-Host "  Interactive Testing: SOME TESTS FAILED" -ForegroundColor Red
    Write-Host "  Review failed tests and fix issues." -ForegroundColor Yellow
    exit 1
}
else {
    Write-Host "  Interactive Testing: ALL SKIPPED" -ForegroundColor Yellow
    Write-Host "  Re-run without -Quick to validate UI/UX." -ForegroundColor Yellow
    exit 0
}
