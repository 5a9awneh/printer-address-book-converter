# Demo CSV Samples

This directory contains **sanitized demo CSV files** for testing and demonstration purposes. All PII (Personally Identifiable Information) has been removed and replaced with mock data.

## 📁 Files

- **`Canon_Sample.csv`** - Canon iR-ADV/imageFORCE format (17 contacts)
- **`Sharp_Sample.csv`** - Sharp MX/BP series format (14 contacts)
- **`Xerox_Sample.csv`** - Xerox AltaLink/VersaLink format (4 contacts)
- **`Develop_Sample.csv`** - Develop/Konica/Bizhub ineo+ format (25 contacts)
- **`Outlook_Sample.txt`** - Outlook "Check Names" format (10 contacts)

**Total: 70 mock contacts** across 4 printer brands + Outlook format

## 🔒 Privacy & Sanitization

All files use mock data from a consistent list of 25 fictional people:
- **Generic names**: John Smith, Jane Doe, Michael Johnson, Emily Davis, Robert Wilson, etc.
- **Example.com emails**: john.smith@example.com, jane.doe@example.com, etc.
- **Placeholder company**: Example Org
- **No real PII**: All names, emails, and organizational references sanitized
- **Encrypted placeholders** preserved for Sharp authentication fields (original format maintained)
- **Generic UUIDs** preserved for Canon (original format maintained)

## ✅ Format Compliance

Each sample file:
- ✅ Maintains **exact CSV structure** of real printer exports
- ✅ Preserves **headers, footers, column counts**
- ✅ Includes **all required fields** per brand specification
- ✅ Uses **correct encoding** (UTF-8 for Canon/Sharp/Xerox, UTF-16LE for Develop)
- ✅ Matches **native printer export format** exactly

## 🔄 Using Your Own Files

The converter allows you to browse/select/enter full paths to your own printer export files. Output files are saved in the same directory as your source file.

```powershell
# Interactive mode with file browser
.\Convert-PrinterAddressBook.ps1

# Or specify paths directly (output saved as: Sharp_Export_to_Canon.csv in same folder)
.\Convert-PrinterAddressBook.ps1 -SourcePath "C:\path\to\Sharp_Export.csv" -TargetPath "C:\path\to\Canon_Template.csv" -NoInteractive
```

No need to place files in any specific directory - the converter works with files from any location.

### Regenerating These Demo Samples

These samples were pre-generated with mock data. If needed for development, they can be regenerated from real exports using custom sanitization scripts (not included in repo for PII protection).

## 🧪 Testing

These files are used by:
- `Test-Comprehensive.ps1` - Full test suite (16 brand-pair conversions)
- `Test-Phase3.ps1` - Phase 3 conversion pipeline tests
- `Test-Interactive.ps1` - Manual workflow validation
- Non-interactive CLI testing: `.\Convert-PrinterAddressBook.ps1 -SourcePath "demo_samples\Sharp_Sample.csv" -TargetPath "demo_samples\Canon_Sample.csv" -NoInteractive`

## 📋 Real Printer Exports

When using the converter with real data, simply browse/select your printer export files from any location on your system.

---

**Safe for public repositories** ✅ No confidential information.
