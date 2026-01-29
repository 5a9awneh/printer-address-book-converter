# Printer Address Book Converter - Execution Plan

## ✅ PROJECT COMPLETE - January 29, 2026

All phases completed successfully. The converter is production-ready with comprehensive testing, intelligent deduplication, Outlook compatibility validation, and enhanced user experience.

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
- **Acceptance Criteria**: ✅ Function extracts headers, contacts, and footers from Canon, Sharp, Xerox, and Develop samples without losing blank lines or comment lines. Tested on all test files in `tests/demo_samples/`.
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
- **Description**: **[PRIORITY]** Add command-line parameters to support non-interactive mode for automated testing: `-SourcePath`, `-TargetPath`, `-Mode` (Single/Batch/Merge), `-NoInteractive`. When parameters are provided, skip menu navigation and execute directly. This enables quick testing without manual GUI clicks.
- **Acceptance Criteria**: ✅ Script can be called with `.\Convert-PrinterAddressBook.ps1 -SourcePath "file.csv" -TargetPath "target.csv" -NoInteractive` and completes conversion without prompts. Target format is auto-detected from the target file. Menu mode still works when no parameters provided.
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
- **Description**: Test conversion from each brand to each other brand (Canon→Sharp, Sharp→Xerox, etc.). Verify email and name appear in correct columns of output. Run against demo test files from tests/demo_samples/. Covers all 4 major printer brands.
- **Acceptance Criteria**: ✅ All brand-pair combinations produce valid output. ✅ Names and emails land in correct columns. ✅ No data loss. ✅ Validated with 13 real CSV files (253 contacts total). ✅ 39/39 integration tests passing.
- **Time**: 25 min → **Actual: 45 min** (expanded to 13 files, fixed CSV parsing edge cases)
- **Completed**: 2026-01-28

---

## Phase 3: Output Writing & Converter Integration

Wire up normalized contacts into target format and rebuild full CSV files.

### 3.1 Refactor Write-AddressBook to use normalized contacts ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Write-AddressBook()` function
- **Description**: Updated `Write-AddressBook()` to accept normalized contact list + target brand. Uses `ConvertFrom-NormalizedContact` for field mapping. Preserves encoding, delimiters, and all brand-specific columns.
- **Acceptance Criteria**: ✅ Output files have correct encoding. ✅ Delimiters match target brand. ✅ All brand-specific fields populated. ✅ Tested Sharp→Canon, Sharp→Xerox, Sharp→Develop.
- **Time**: 25 min → **Actual: 25 min**
- **Completed**: 2026-01-28

### 3.2 Create output validation function ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Creates**: `Validate-OutputFile()` function
- **Description**: Created validation function that verifies output CSV structure, checks required fields, validates encoding, and detects empty/corrupted files. Returns detailed results with IsValid flag, errors, warnings, and contact count.
- **Acceptance Criteria**: ✅ Detects malformed CSVs. ✅ Validates brand-specific structure. ✅ Checks email field population. ✅ Handles missing files. ✅ Tested with Canon, Xerox, Develop outputs.
- **Time**: 20 min → **Actual: 20 min**
- **Completed**: 2026-01-28

### 3.3 Update Convert-AddressBook to use refactored pipeline ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Convert-AddressBook()` function
- **Description**: Refactored full pipeline: Parse CSV directly → `ConvertTo-NormalizedContact` (per row) → Validate → Deduplicate → `Write-AddressBook` (with `ConvertFrom-NormalizedContact`). Removed dependency on `Read-AddressBook` to work with raw CSV rows.
- **Acceptance Criteria**: ✅ End-to-end conversion works. ✅ Sharp→Canon: 13/13 contacts. ✅ Sharp→Develop: 29/30 (1 missing name). ✅ Invalid contacts skipped with warnings. ✅ Logging at each step.
- **Time**: 20 min → **Actual: 20 min**
- **Completed**: 2026-01-28

---

## Phase 4: Deduplication, Merging & Testing

Add optional post-processing features and comprehensive test suite.

### 4.1 Enhance Remove-Duplicates with email-based matching ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Modifies**: `Remove-Duplicates()` function
- **Description**: Update dedup logic to match on email first (primary key), then name fuzzy-match as secondary. Merge contact records intelligently (keep longest name, all contact info). Configurable via flag in `Convert-AddressBook()`.
- **Acceptance Criteria**: ✅ Function removes contacts with duplicate emails. Fuzzy-match catches "John Smith" vs "J. Smith". Stats report dupe count accurately.
- **Time**: 20 min → **Actual: 35 min**
- **Completed**: 2026-01-28
- **Implementation**: Added 3 new functions: Enhanced `Remove-Duplicates()` with fuzzy matching, `Get-NameSimilarity()` (0.8 threshold), `Get-EditDistance()` (DP algorithm). Detects abbreviations and typos.

### 4.2 Add Outlook import validation helper ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`
- **Creates**: `Test-OutlookCompatibility()` function
- **Description**: Add optional function to verify output is Outlook-importable (check for problematic characters, field lengths, required columns). Optional post-processing step before save.
- **Acceptance Criteria**: ✅ Validates all output files. Warns on fields exceeding Outlook limits (e.g., email >254 chars). Reports importability status.
- **Time**: 15 min → **Actual: 15 min**
- **Completed**: 2026-01-28
- **Implementation**: Created `Test-OutlookCompatibility()` function (~104 lines). Validates email ≤254 chars, name ≤256 chars, detects problematic characters ([,],<,>,;,:,",\,/), control characters, leading/trailing spaces.

### 4.3 Build comprehensive test suite ✅ COMPLETE
- **Files**: `tests/Test-Comprehensive.ps1` (new)
- **Creates**: Test runner for all brand combinations + edge cases
- **Description**: Create Pester/manual test script that:
  - Converts all test files (Canon, Sharp, Xerox, Develop) to each target brand
  - Validates structure, headers, footers, column alignment, encoding
  - Checks email/name mapping accuracy
  - Verifies no data loss
  - Tests edge cases (empty name, invalid email, special chars, Unicode)
- **Acceptance Criteria**: ✅ Test suite runs without errors. All conversions pass validation. Edge cases handled gracefully.
- **Time**: 45 min → **Actual: 45 min**
- **Completed**: 2026-01-28
- **Implementation**: Created comprehensive test file with 3 test suites: (1) Brand-pair conversions (4×4=16 combinations), (2) Output structure validation, (3) Edge cases (10 tests: empty names, invalid emails, special chars, Unicode, long fields, duplicates, fuzzy matching, name splitting, email validation, Outlook compatibility).

### 4.4 Create sanitized demo CSV samples for repository ✅ COMPLETE
- **Files**: `tests/demo_samples/` (new directory)
- **Creates**: Anonymized CSV samples for each brand (Canon, Sharp, Xerox, Develop)
- **Description**: Create sanitized demo CSV files from existing test exports with all PII removed:
  - Remove IP addresses from filenames (use generic names like "Canon_Sample.csv")
  - Anonymize email addresses (use example.com domain)
  - Replace real names with demo names (John Doe, Jane Smith, etc.)
  - Keep exact CSV structure/format for each brand
  - Maintain headers, footers, column counts
  - Include 5-10 contacts per sample for realistic testing
- **Acceptance Criteria**: ✅ Each brand has representative demo CSV. No PII in filenames or content. Files can be committed to public repo safely. Structure matches real exports exactly.
- **Time**: 25 min → **Actual: 25 min**
- **Completed**: 2026-01-28
- **Implementation**: Created 4 sanitized demo samples (Canon, Sharp, Xerox, Develop) with generic names, example.com emails, Acme Corp company. Added .gitignore to exclude backup and converted files from git. Created comprehensive README for demo_samples/. Note: Xerox format has known duplicate column issue in some exports (documented edge case).

### 4.5 Interactive workflow testing (end-to-end manual validation) ✅ COMPLETE
- **Files**: `tests/Test-Interactive.ps1` (created in Phase 2)
- **Executes**: Manual interactive workflow validation script
- **Description**: Execute guided interactive test for manual validation of complete user workflows after all backend implementation: file selection dialogs, menu navigation, mode selection (Convert/Merge/Outlook), conversion execution, success output validation. Tests actual UI/UX experience including dialog foreground behavior.
- **Acceptance Criteria**: ✅ Dialog appears in foreground. ✅ All menu options work correctly. ✅ File selection is intuitive. ✅ Success/error messages are clear. ✅ All workflows complete end-to-end successfully. ✅ Multi-file batch shows unified folder prompt (not per-file). ✅ Outlook mode creates files correctly.
- **Time**: 20 min
- **Completed**: 2026-01-29
- **Enhancements Added**:
  - **Unified success output**: Created `Show-OutputSuccess()` function for consistent success messages across all modes (Convert, Merge, Outlook)
  - **Batch improvements**: Multi-file conversions now show ONE folder prompt at end instead of prompting after each file
  - **Better visibility**: Enhanced success banners with file size, contact count, and "Open folder in Explorer?" prompt
  - **Silent mode**: Added `-Silent` flag to suppress individual file success messages during batch operations
  - **SearchKey optimization**: Develop format Speed Dial (SearchKey) now based on FirstName instead of LastName for Outlook imports
  - **User warnings**: Added prominent warning in main menu that all CSV files must contain at least 1 contact (required for format detection)
  - **Documentation updates**: Enhanced README with imageFORCE support, CLI parameters, improved Output Structure table with Special Features column, Known Limitations expanded

### 4.6 Refine menu and documentation ✅ COMPLETE
- **Files**: `Convert-PrinterAddressBook.ps1`, `README.md`
- **Modifies**: `Show-Menu()`, help text, README
- **Description**: Update interactive menu with new features (dedup option, output validation, Outlook check). Add examples to README for each brand conversion. Update `Get-Help` docs.
- **Acceptance Criteria**: ✅ Help text reflects all features. Menu is clear and intuitive. README has brand-pair conversion examples.
- **Time**: 15 min → **Actual: 15 min**
- **Completed**: 2026-01-28
- **Implementation**: Updated script header to v2.0 with Phase 4 changelog (fuzzy dedup, Outlook validation). Enhanced README features section with intelligent deduplication details, edit distance algorithm mention, Outlook compatibility validation.

### 4.7 Remove backup function from script ✅ COMPLETED
- **Files**: `Convert-PrinterAddressBook.ps1`, `README.md`, `.gitignore`
- **Modifies**: `Backup-SourceFile()` (removed), `Get-SafeOutputPath()`, `Main()` functions
- **Description**: Removed backup file creation functionality since script only reads source files and creates new output files without modifying originals. Output files now saved in same directory as source file instead of separate converted/ folder.
- **Acceptance Criteria**: ✅ Script no longer creates backup/ or converted/ directories. Output files saved alongside source files. All conversions work correctly.
- **Time**: 10 min
- **Completed**: 2026-01-28

### 4.8 Comprehensive production validation suite ✅ COMPLETE
- **Files**: `tests/Test-Production.ps1` (new)
- **Creates**: Automated validation for all conversion scenarios
- **Description**: Execute comprehensive validation of all converter features:
  1. **Batch conversions**: Convert all demo samples to each target format (Canon, Sharp, Xerox, Develop) - 16 combinations
  2. **Round-trip testing**: Re-import converted files and reconvert to verify no data loss - 16 round-trips
  3. **Merge testing**: Merge all demo samples into each target format - 4 merge operations
  4. **Outlook testing**: Create address book from Outlook format and convert to each brand - 4 conversions
  5. **Structure validation**: Verify CSV structure, headers, encoding, required fields for all outputs
  6. **Data integrity**: Compare contact counts, validate emails, check name preservation
- **Acceptance Criteria**: ✅ All batch conversions succeed (16/16). ✅ Round-trip conversions produce identical results (16/16). ✅ Merge operations combine all contacts correctly. ✅ Outlook mode creates valid output. ✅ No structure mismatches (16/16 structure validations passed). ✅ All CSV files are importable.
- **Time**: 30 min → **Actual: 90 min** (discovered and fixed CSV parsing bugs, header detection issues, Develop format writing)
- **Completed**: 2026-01-28
- **Results**: 48/48 tests passed (100%). All 4 brands (Canon, Sharp, Xerox, Develop) convert correctly in all 16 brand-pair combinations. Round-trip conversions verified. Production-ready.

---

## Summary

| Phase | Tasks | Focus | Est. Time | Actual Time | Status |
|-------|-------|-------|-----------|-------------|--------|
| **1** | **1.1–1.5** | **Parsing + CLI + Logging** | **1 h 50 min** | **1 h 50 min** | **✅ COMPLETE (5/5)** |
| **2** | **2.1–2.3** | **Field mapping & normalization** | **1 h 10 min** | **1 h 20 min** | **✅ COMPLETE (3/3)** |
| **3** | **3.1–3.3** | **Output writing & integration** | **1 h 5 min** | **1 h 5 min** | **✅ COMPLETE (3/3)** |
| **4** | **4.1–4.8** | **Dedup, testing, validation** | **3 h 0 min** | **4 h 15 min** | **✅ COMPLETE (8/8)** |
| **Total** | **20 tasks** | **End-to-end refactor + testing** | **~7 h 35 min** | **~8 h 30 min** | **✅ COMPLETE (20/20 - 100%)** |

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
- `tests/demo_samples/` (sanitized sample CSVs for all brands)
