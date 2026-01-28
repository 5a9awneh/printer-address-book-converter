# Printer Address Book Converter - Execution Plan (planx.md)

## Overview

This project converts printer address book CSV exports between different brands (Canon, Sharp, Xerox, Develop/Konica, Bizhub) using a **brand-agnostic approach**:

- **No proprietary logic per brand**: Only configuration maps brand-specific headers to normalized fields
- **Header/Footer preservation**: Extract and re-inject source headers/footers unchanged to maintain CSV validity
- **Contact data isolation**: Identify contact rows (lines with `@` symbol for emails) and map email + name between formats
- **Minimal transformation**: Extract email + full name from source, validate, normalize (optional), and write to target format
- **Validators stay**: Keep email syntax, empty field, and name validation rules
- **Nice-to-haves**: Deduplication and normalization applied post-conversion

The conversion pipeline: **Read (auto-detect) → Parse (headers/data/footers) → Map (normalize) → Validate → Write**

---

## Phase 1: Core Parsing Refactor

Extract header/footer logic and isolate contact data blocks for all brands.

### 1.1 Refactor header/footer detection into reusable function ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Creates**: `Extract-CsvStructure()` function
- **Description**: Build a function that analyzes CSV files and separates headers (lines before first contact), contact data (rows with `@`), and footers (lines after last contact). This standardizes parsing across all brands and preserves empty lines—critical for CSV validity. Returns object with `@{Headers=[], Contacts=[], Footers=[]}`.
- **Acceptance Criteria**: ✅ Function extracts headers, contacts, and footers from Canon, Sharp, Xerox, and Develop samples without losing blank lines or comment lines. Tested on all test files in `tests/source_exports/`.
- **Time**: 30 min → **Actual: 30 min**
- **Completed**: 2026-01-28

### 1.2 Update Read-AddressBook to use new extraction function ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Read-AddressBook()` function
- **Description**: Replace manual parsing logic with calls to `Extract-CsvStructure()`. Remove brand-specific header/footer parsing loops and use the new unified approach. Ensures all brands are parsed consistently.
- **Acceptance Criteria**: ✅ `Read-AddressBook()` calls `Extract-CsvStructure()` and builds contact objects with email and name fields. Test with one Canon, one Sharp file.
- **Time**: 20 min → **Actual: 20 min**
- **Completed**: 2026-01-28

### 1.3 Add CLI parameter support for non-interactive testing ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `param()` block and `Main` function
- **Description**: **[PRIORITY]** Add command-line parameters to support non-interactive mode for automated testing: `-SourcePath`, `-TargetBrand`, `-Mode` (Single/Batch/Merge), `-NoInteractive`. When parameters are provided, skip menu navigation and execute directly. This enables quick testing without manual GUI clicks.
- **Acceptance Criteria**: ✅ Script can be called with `.\Convert-PrinterAddressBook.ps1 -SourcePath "file.csv" -TargetBrand "Canon" -NoInteractive` and completes conversion without prompts. Menu mode still works when no parameters provided.
- **Time**: 20 min → **Actual: 20 min**
- **Completed**: 2026-01-28

### 1.4 Add tests for header/footer preservation ✅ COMPLETE
- **Files**: `tests/` (create new test file or expand existing)
- **Creates**: Test cases for `Extract-CsvStructure()`
- **Description**: Write simple tests that verify headers and footers (including blank lines) are preserved exactly. Parse a sample file, reconstruct it, and compare line-by-line. Uses new CLI parameter support for automated testing.
- **Acceptance Criteria**: ✅ All test files parse and reconstruct without line loss. Empty lines in headers/footers remain intact.
- **Time**: 15 min → **Actual: 20 min** (created comprehensive test suite with 3 suites, 29 tests)
- **Completed**: 2026-01-28

### 1.5 Enhance logging for development debugging ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Write-Log()` function and add logging throughout pipeline
- **Description**: **[PRIORITY]** Enhance existing logging to include:
  - Function entry/exit with parameters (DEBUG level)
  - Data transformation steps with before/after counts
  - Detailed error messages with stack traces
  - Add `-Verbose` parameter support to show logs in console
  - Log file location printed at script start
  This enables better debugging/tracing during development.
- **Acceptance Criteria**: ✅ All major functions log entry/exit. ✅ Errors include full context. ✅ `-Verbose` shows real-time logging. ✅ Test runs show clear trace of execution flow.
- **Time**: 25 min → **Actual: 20 min**
- **Completed**: 2026-01-28

---

## Phase 2: Field Mapping Engine

Build normalized mapping layer for email and name fields.

### 2.1 Create unified contact schema and normalization functions ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Creates**: `Normalize-Contact()` function (email validation, name sanitization, empty-field checks)
- **Description**: Define a standard contact object schema `@{Email, FirstName, LastName, DisplayName}`. Build `Normalize-Contact()` to accept source contact, brand name, and target brand—then map email + name between formats. Apply existing validators (`Test-Email()`, name length checks) during normalization.
- **Acceptance Criteria**: ✅ Function maps Canon email + name to Sharp format and vice versa. ✅ Rejects contacts with invalid emails or missing names. ✅ Handles multi-word names correctly.
- **Time**: 30 min → **Actual: 25 min**
- **Completed**: 2026-01-28

### 2.2 Extend BrandConfig to define field mappings ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `$Script:BrandConfig`
- **Description**: Add mapping rules to each brand config: `EmailField`, `NameField`, `FirstNameField`, `LastNameField` (if applicable). Ensure every brand has clear target field names for output. This decouples field names from business logic.
- **Acceptance Criteria**: ✅ All four brands (Canon, Sharp, Xerox, Develop) have complete field mappings. ✅ Verify against test CSVs.
- **Time**: 15 min → **Actual: 10 min**
- **Completed**: 2026-01-28

### 2.3 Test field mapping with multi-brand conversions ✅ COMPLETE
- **Files**: `tests/Test-Phase2-Integration.ps1`
- **Creates**: Integration test suite for all brand-pair conversions
- **Description**: Test conversion from each brand to each other brand (Canon→Sharp, Sharp→Xerox, etc.). Verify email and name appear in correct columns of output. Run against real test files from tests/source_exports/. Expanded to test ALL 13 CSV files.
- **Acceptance Criteria**: ✅ All brand-pair combinations produce valid output. ✅ Names and emails land in correct columns. ✅ No data loss. ✅ Validated with 13 real CSV files (253 contacts total). ✅ 39/39 integration tests passing.
- **Time**: 25 min → **Actual: 45 min** (expanded to 13 files, fixed CSV parsing edge cases)
- **Completed**: 2026-01-28

### 2.4 Interactive workflow testing (moved from Phase 4.4)
- **Files**: `tests/Test-Interactive.ps1` (new)
- **Creates**: Manual interactive workflow validation script
- **Description**: Create guided interactive test for manual validation of user workflows: file selection dialogs, menu navigation, mode selection (Single/Batch/Merge), conversion execution, backup/restore validation. Tests actual UI/UX experience including dialog foreground behavior.
- **Acceptance Criteria**: Dialog appears in foreground. All menu options work correctly. File selection is intuitive. Success/error messages are clear. Backup/restore functions properly.
- **Time**: 20 min
- **Note**: Moved earlier due to dialog display issues during manual testing

---

## Phase 3: Output Writing & Converter Integration

Wire up normalized contacts into target format and rebuild full CSV files.

### 3.1 Refactor Write-AddressBook to use normalized contacts
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Write-AddressBook()` function
- **Description**: Update `Write-AddressBook()` to accept normalized contact list + target brand. Reconstruct full CSV by injecting contacts between headers and footers. Preserve encoding, delimiters, and all columns (fill unmapped fields with defaults/blanks).
- **Acceptance Criteria**: Output files have correct encoding (UTF8 for Canon/Sharp/Xerox, Unicode for Develop). Delimiters match target brand. No column misalignment.
- **Time**: 25 min

### 3.2 Create output validation function
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Creates**: `Validate-OutputFile()` function
- **Description**: Verify output CSV has same structure as source (same number of columns, correct header/footer format). Check that no rows are missing and that email/name fields are populated in correct columns. Can be used pre-save as quality gate.
- **Acceptance Criteria**: Function detects malformed CSVs (wrong column count, missing headers, corrupted footers). Passes on all good outputs.
- **Time**: 20 min

### 3.3 Update Convert-AddressBook to use refactored pipeline
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Convert-AddressBook()` function
- **Description**: Wire together: `Detect-Brand()` → `Read-AddressBook()` → `Extract-CsvStructure()` → `Normalize-Contact()` → `Validate-Contacts()` → `Write-AddressBook()` → `Validate-OutputFile()`. Each step feeds the next. Log results.
- **Acceptance Criteria**: End-to-end conversion works for all brand combinations. Invalid contacts are skipped with warning. Output is valid and importable.
- **Time**: 20 min

---

## Phase 4: Deduplication, Merging & Testing

Add optional post-processing features and comprehensive test suite.

### 4.1 Enhance Remove-Duplicates with email-based matching
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Remove-Duplicates()` function
- **Description**: Update dedup logic to match on email first (primary key), then name fuzzy-match as secondary. Merge contact records intelligently (keep longest name, all contact info). Configurable via flag in `Convert-AddressBook()`.
- **Acceptance Criteria**: Function removes contacts with duplicate emails. Fuzzy-match catches "John Smith" vs "J. Smith". Stats report dupe count accurately.
- **Time**: 20 min

### 4.2 Add Outlook import validation helper
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Creates**: `Export-OutlookCompatible()` or validation function
- **Description**: Add optional function to verify output is Outlook-importable (check for problematic characters, field lengths, required columns). Optional post-processing step before save.
- **Acceptance Criteria**: Validates all output files. Warns on fields exceeding Outlook limits (e.g., email >254 chars). Reports importability status.
- **Time**: 15 min

### 4.3 Build comprehensive test suite
- **Files**: `tests/test-suite.ps1` (new)
- **Creates**: Test runner for all brand combinations + edge cases
- **Description**: Create Pester/manual test script that:
  - Converts all test files (Canon, Sharp, Xerox, Develop) to each target brand
  - Validates structure, headers, footers, column alignment, encoding
  - Checks email/name mapping accuracy
  - Verifies no data loss
  - Tests edge cases (empty name, invalid email, special chars, Unicode)
- **Acceptance Criteria**: Test suite runs without errors. All conversions pass validation. Edge cases handled gracefully.
- **Time**: 45 min

### 4.4 Create integration test with user workflows ✅ MOVED TO 2.4
- **Status**: Moved to Task 2.4 (Phase 2) due to immediate need for UI validation
- **See**: Task 2.4 for implementation details

### 4.5 Refine menu and documentation
- **Files**: `Convert-PrinterAddressBook.ps1`, `README.md`
- **Modifies**: `Show-Menu()`, help text, README
- **Description**: Update interactive menu with new features (dedup option, output validation, Outlook check). Add examples to README for each brand conversion. Update `Get-Help` docs.
- **Acceptance Criteria**: Help text reflects all features. Menu is clear and intuitive. README has brand-pair conversion examples.
- **Time**: 15 min

---

## Summary

| Phase | Tasks | Focus | Est. Time | Actual Time | Status |
|-------|-------|-------|-----------|-------------|--------|
| **1** | **1.1–1.5** | **Parsing + CLI + Logging** | **1 h 50 min** | **1 h 50 min** | **✅ COMPLETE** |
| **2** | **2.1–2.4** | **Field mapping + normalization + interactive tests** | **1 h 30 min** | **1 h 20 min** | **⏳ 3/4 tasks** |
| 3     | 3.1–3.3 | Output writing & integration | 1 h 5 min | — | Not started |
| 4     | 4.1–4.5 | Dedup, Outlook, testing, docs | 1 h 55 min | — | Not started |
| **Total** | **18 tasks** | **End-to-end refactor + testing** | **~6 h** | **~3 h 10 min** | **8/18 tasks (44%)** |

---

## How to Use This Plan

1. **Pick a task**: e.g., "Execute task 1.1 from planx.md"
2. **Follow the title and description**: Implement the function/modification described
3. **Verify acceptance criteria**: Test the specific scenarios listed
4. **Move to next task**: Tasks build on each other within phases, but phases can run in order
5. **Track progress**: Update task status in your workflow (e.g., ✓ Done, ⏳ In Progress)

Each task is sized to be completable in one focused session (15–45 min). Parallel execution possible within Phase 4.

### Documentation Guidelines

- **Keep it minimal**: Update ONLY this plan file with task completion status
- **No separate reports**: Unless explicitly requested, avoid creating validation reports or summary documents
- **Surgical updates**: Mark tasks as ✅ COMPLETE with actual time and completion date inline
- **Commit messages**: Use detailed git commit messages to document what was done
- **Let git tell the story**: Code changes and commit history are the primary documentation

### Progress Tracking Format

When completing a task, update the task header and add completion metadata:
```
### X.X Task Name ✅ COMPLETE
- **Acceptance Criteria**: ✅ [check each criterion]
- **Time**: [estimate] → **Actual: [actual time]**
- **Completed**: YYYY-MM-DD
```

---

## Key Symbols Reference

**Functions to create:**
- `Extract-CsvStructure()`
- `Normalize-Contact()`
- `Validate-OutputFile()`
- `Export-OutlookCompatible()` (optional)

**Functions to modify:**
- `Read-AddressBook()`
- `Write-AddressBook()`
- `Convert-AddressBook()`
- `Remove-Duplicates()`
- `Show-Menu()`

**Config to extend:**
- `$Script:BrandConfig` (add field mappings)
- `$Script:Stats` (add Outlook/dedup stats if needed)

**Test files location:**
- `tests/source_exports/` (sample CSVs for all brands)
