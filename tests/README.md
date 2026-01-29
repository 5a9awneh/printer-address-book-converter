# Test Suite

Comprehensive testing infrastructure for the Printer Address Book Converter.

## 🔄 Demo Samples

The `demo_samples/` directory contains pre-generated sanitized CSV files ready for testing.

**These samples:**
- Use mock data only (no real PII)
- Maintain exact structural fidelity to real printer exports
- Are safe for public repositories and distribution
- Work with all converter features and test suites

**Mock Data:**
- Names: John Smith, Jane Doe, Michael Johnson, Emily Davis, etc.
- Emails: name@example.com
- Organization: Example Org

To regenerate samples from your own printer exports, edit `Sanitize-Samples.ps1` to point to your source files.

### Verify Sample Integrity

```powershell
cd tests
.\Verify-MockStructure.ps1
```

**Validates:**
- ✅ Proper CSV structure
- ✅ Headers/footers present
- ✅ Column structure intact
- ✅ Field counts correct

## 🧪 Running Tests

### Full Test Suite ⭐
```powershell
.\Test-All.ps1
```
Comprehensive automated testing covering:
- 16 brand-pair conversions
- 4 merge operations
- 16 round-trip tests
- 1 Outlook import test
- **Total: 37 tests**

**Options:**
```powershell
.\Test-All.ps1              # Auto cleanup after tests
.\Test-All.ps1 -KeepOutputs # Keep generated files for inspection
```

### Single Conversion Test (CLI)
```powershell
cd ..
.\Convert-PrinterAddressBook.ps1 -SourcePath "tests\demo_samples\Sharp_Sample.csv" -TargetPath "tests\demo_samples\Canon_Sample.csv" -NoInteractive
```
*Note: TargetPath is used only for format detection*

## 📊 Test Coverage

- **4 Printer Brands**: Canon, Sharp, Xerox, Develop/Konica Minolta
- **16 Conversion Pairs**: All brand-to-brand combinations
- **60 Total Contacts**: Across all demo samples
  - Canon: 17 contacts
  - Sharp: 14 contacts
  - Xerox: 4 contacts
  - Develop: 25 contacts

## 🔒 Privacy & Security

- ✅ `demo_samples/` - Safe for public repositories (mock data only)
- 🔐 Users provide their own source files when using the converter
- 📁 Converter supports browsing/selecting files from any location

## 📝 Notes

- All demo samples maintain **exact structural fidelity** to real printer exports
- Original encodings preserved (UTF-8 for Canon/Sharp/Xerox, UTF-16LE for Develop)
- Encrypted authentication fields maintained in Sharp format
- UUIDs preserved in Canon format
