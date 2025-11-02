
# CRM Encoding Repair

**Advanced VBA script for fixing UTF-8/CP-1252 encoding issues in CRM exports**
**Renamed project** from `bogus-character-cleanup` to `crm-encoding-repair` to better reflect purpose and target audience

## The Problem

When working with international contact data in CRMs (Salesforce, HubSpot, Dynamics, etc.), encoding issues cause names and text to become corrupted during export/import:

- `Müller` → `MÃ¼ller`
- `François` → `Françoise`
- `Święty` → `ÅšwiÄ™ty`
- `Malmö` → `Malmö`

This happens when UTF-8 encoded data is incorrectly read as CP-1252/Windows-1252, making the data unusable for database uploads, mail merges, and automation.

## The Solution

This VBA macro automatically detects and fixes 250+ common encoding corruption patterns in a single pass, with detailed logging of every change made.

## Features

- ✅ **250+ encoding patterns** covering Western, Eastern, and Nordic European languages
- ✅ **Comprehensive CP-1252 support** including smart quotes, dashes, currency symbols
- ✅ **Detailed audit logging** - timestamped log file with before/after values for every change
- ✅ **70% faster** than naive sequential checking
- ✅ **Smart processing** - only updates cells that need fixing
- ✅ **Triple-encoding detection** for data that's been corrupted multiple times

## Language Support

**Western European:** German, French, Spanish, Portuguese, Italian, Dutch  
**Nordic:** Swedish, Norwegian, Danish, Icelandic, Finnish  
**Eastern European:** Polish, Czech, Slovak, Romanian, Hungarian, Croatian  
**Baltic:** Latvian, Lithuanian, Estonian  
**Other:** Turkish, Albanian

## Installation

1. Open Excel with your corrupted CRM export
2. Press `Alt + F11` to open VBA Editor
3. Insert → Module
4. Paste the contents of `FixEncodingWithLogging.vba`
5. Save and close VBA Editor

## Usage

1. Open your spreadsheet with encoding issues
2. Activate the worksheet you want to fix
3. Press `Alt + F8` to open Macros dialog
4. Select `FixingEncodingIssuesAdvanced`
5. Click **Run**

The script will:
- Process all cells in the active worksheet
- Fix all detected encoding issues
- Create a timestamped log file (e.g., `EncodingFixes_2025-11-01_14-30-45.txt`)
- Display a summary of changes

## Documentation

- **[Full Documentation](EncodingFix_Documentation.md)** - Complete guide with troubleshooting
- **[Quick Reference Guide](Quick_Reference_Guide.md)** - Common issues by language
- **[Sample Log File](Sample_EncodingFixes_Log.txt)** - Example output

## What Gets Fixed

### Common Patterns
- **Accented letters:** Ã©→é, Ã¼→ü, Ã±→ñ, Å‚→ł
- **Smart quotes:** â€œ→", â€™→'
- **Dashes:** â€"→–, â€"→—
- **Currency:** â‚¬→€, Â£→£
- **Symbols:** Â©→©, Â®→®, Â°→°
- **HTML entities:** &#39;→', &quot;→"

### Example Names Fixed
- MÃ¼ller → Müller
- Françoise → Françoise
- ÅšwiÄ™ty → Święty
- Malmö → Malmö
- Jiří Dvořák → Jiří Dvořák
- O&#39;Brien → O'Brien

## Common Use Cases

- ✅ Salesforce contact exports
- ✅ HubSpot data migrations
- ✅ Microsoft Dynamics 365 reports
- ✅ Zoho CRM exports
- ✅ Pipedrive CSV downloads
- ✅ Any CSV/Excel export with international names

## Technical Details

This script fixes the specific issue where UTF-8 encoded text (the modern standard) is incorrectly interpreted as CP-1252/Windows-1252 encoding. This commonly occurs when:

- Exporting from web-based CRMs to Excel
- Importing CSV files with wrong encoding
- Migrating data between systems
- Copy/pasting from emails or web forms

The script uses a dictionary-based approach for O(1) lookup performance, processing each cell only once.

## Version History

### v2.0 - Advanced (November 2025)
- Added logging system with timestamped audit files
- Expanded from 40 to 250+ encoding patterns
- Added CP-1252 complete coverage
- Added support for Czech, Slovak, Turkish, Romanian, Baltic languages
- Improved performance (70% faster)
- Added triple-encoding detection

### v1.0 - Basic (Original)
- Core UTF-8 encoding fixes for German, French, Spanish, Polish
- Manual if/then checking
- ~40 common patterns

## Contributing

Found an encoding pattern not covered? Submit an issue with:
1. The corrupted text (e.g., "MÃ¼ller")
2. What it should be (e.g., "Müller")
3. The language/origin

## License

Free to use and modify for your organization. Attribution appreciated but not required.

---

**Party on Wayne!** 🎸
