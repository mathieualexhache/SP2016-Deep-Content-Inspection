# SharePoint 2016 Document Security Scanner

A comprehensive PowerShell-based toolkit for scanning and auditing documents stored in SharePoint 2016 on-premise libraries. This suite performs static analysis to detect password protection, MIME mismatches, embedded macros (VBA, etc.), and file encryption using a combination of native PowerShell, external tools (7-Zip, Siegfried), and Python-based macro detection.

---

## Repository Contents

| Script/File                     | Purpose                                                                 |
|--------------------------------|-------------------------------------------------------------------------|
| `check_pwd_protected_encrypted` | Scans for password-protected Office, ZIP, and PDF files using static headers and 7-Zip |
| `renamed_extension` | Compares expected MIME types (based on file extension) with actual content signatures using Siegfried |
| `macro_scan`        | Detects embedded macros in Office files using a Python-based scanner |
| `OleMacroDetector.py`          | Python macro triage engine combining olevba, oledump, VBA_Parser, and ZIP/OLE inspection |

---

## Features

### 1. **Password Protection Detection**
- Detects encrypted Office 97–2003 and Office 2007+ formats
- Scans ZIP archives using 7-Zip CLI
- Identifies encrypted PDFs via static markers
- Outputs timestamped CSV and log files

### 2. **MIME Type Validation**
- Uses Siegfried (`sf.exe`) with custom extension-agnostic signature
- Compares actual MIME type with expected type from file extension
- Logs unknown formats and mismatches
- Outputs CSV report and unknowns log

### 3. **Macro Detection**
- Scans Office files for embedded macros using static analysis
- Integrates with Python script `OleMacroDetector.py`
- Supports formats: `.doc`, `.xls`, `.ppt`, `.docm`, `.xlsm`, `.pptm`, `.sldm`, `.dotm`, `.potm`, `.xlsx`, `.pub`, `.mht`, `.slk`, and more
- Outputs CSV report and detailed log

### 4. **Python Macro Triage Engine**
- Combines:
  - `olevba` triage scan
  - `oledump.py` macro stream detection
  - `VBA_Parser` deep inspection
  - Excel 97 OLE stream checks
  - ZIP-based embedded macro scan
- Returns `MACROS_PRESENT` or `NO_MACROS`

---

## Requirements

### General
- Windows workstation (domain-joined)
- PowerShell 5.1 or later
- SharePoint 2016 CSOM DLLs
- Internet access for downloading tools/signatures

### Password Protection Scanner
- [7-Zip](https://www.7-zip.org/) CLI (`7z.exe`) in system path

### MIME Scanner
- [Siegfried](https://github.com/richardlehane/siegfried) (`sf.exe`, `roy.exe`)
- Signature files from The UK National Archives

### Macro Scanner
- [Python](https://www.python.org/) 3.x in system path
- Python packages: `olefile`, `oletools`
- `oledump.py` (auto-downloaded from GitHub)

---

## Getting Started

### 1. Clone the Repository
```bash
git clone https://github.com/your-org/sharepoint-security-scanner-suite.git
cd SP2016-Deep-Content-Inspection
```
## References

Here are the key tools and resources used in this project:

- **[Siegfried](https://github.com/richardlehane/siegfried)** – File format identification tool
- **[oletools](https://github.com/decalage2/oletools)** – Python tools to analyze OLE and MS Office files
- **[oledump.py](https://github.com/DidierStevens/DidierStevensSuite)** – Tool to analyze OLE files and detect macros
- **[PRONOM](https://www.nationalarchives.gov.uk/PRONOM/Default.aspx)** – Technical registry of file formats maintained by The UK National Archives
- **[7-Zip](https://www.7-zip.org/)** – Open-source file archiver used for encryption detection
---

## License

This project is licensed under the **MIT License**. You are free to use, modify, and distribute this software with attribution.
