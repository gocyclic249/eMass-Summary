# eMass-Summary

POA&M Analysis Tool - Processes POA&M Excel files and generates ICS cybersecurity assessment reports using the Ask Sage GenAI API.

> **Note:** The script itself is UNCLASSIFIED, but once eMass data is added, the output becomes CUI. Only CUI-capable LLMs should be used.

## Overview

This PowerShell script extracts POA&M data from Excel files exported from eMass Executive Reports, analyzes each record using the Ask Sage GenAI API, and generates a formatted Word document report with proper CUI markings.

## Features

- Interactive file picker dialogs for input/output selection
- Automatic extraction of essential POA&M fields
- AI-powered analysis of each POA&M record
- Severity mapping (High→CAT I, Moderate→CAT II, Low→CAT III)
- Automatic overdue calculation
- Legacy/unpatchable system detection
- Remediation owner assignment logic
- Risk priority classification (Critical, High, Medium, Low)
- Color-coded Word document output
- CUI markings in header and footer
- Batch processing with rate limiting

## Prerequisites

- Microsoft Excel installed
- Microsoft Word installed
- Ask Sage API key (see setup instructions below)
- PowerShell 5.1 or higher
- eMass Executive Report POA&Ms Export (Excel format)

## Setup Instructions

### 1. Obtain API Key

1. Navigate to Settings in [Ask Sage](https://chat.genai.army.mil)
2. Switch to the **Account** tab
3. Scroll down to "Manage your API Keys" in the sidebar
4. Click to generate your new API key

> **Security:** Keep your API key secure - treat it like a password and rotate regularly.

### 2. Set Environment Variable

Run this command in PowerShell (replace `YOUR_API_KEY` with your actual key):

```powershell
[System.Environment]::SetEnvironmentVariable("GENAI_API_TOKEN", "YOUR_API_KEY", "User")
```

Then restart PowerShell for the change to take effect.

### 3. Verify Setup

Test that the variable is set:

```powershell
[System.Environment]::GetEnvironmentVariable("GENAI_API_TOKEN", "User")
```

## Usage

1. Run the script:
   ```powershell
   .\POAMAnalysis.ps1
   ```
2. Select your Excel POA&M file when prompted
3. Choose output location for the Word document
4. Wait for processing to complete
5. Review the generated `.docx` output

## Excel File Format

The script expects an eMass Executive Report POA&M export with:

- **First two rows:** Ignored (removed automatically)
- **Third row:** Headers
- **Data rows:** POA&M records

### Required Headers

| Field | Description |
|-------|-------------|
| Organization | Base/organization name |
| System Name | Name of the system |
| ID | 13-digit POA&M identifier |
| POA&M URL | Link to the POA&M in eMass |
| Controls / APs | NIST control identifiers (e.g., CM-7.1) |
| POA&M Item Status | Current status |
| Scheduled Completion Date | Target completion date |
| Completion Date | Actual completion date |
| Risk Accepted Date | Date risk was accepted |
| Vulnerability Description | Description of the vulnerability |
| Raw Severity | Original severity value |
| Severity | Mapped severity (High/Moderate/Low) |
| Mitigations | Applied mitigations |
| Recommendations | Recommended actions |

## Output Report

The generated Word document includes:

- **Header:** CUI marking, report title, generation date
- **Body:** Analyzed POA&M records with:
  - POA&M ID and URL
  - Organization and System information
  - Severity (CAT I/II/III) with color coding
  - Status and scheduled completion
  - Days overdue calculation
  - Control families
  - Legacy/unpatchable flag
  - Risk priority (Critical/High/Medium/Low)
  - Vulnerability summary
  - ICS-safe recommended fix
  - Remediation owner assignment
  - Escalation requirements
  - Compensating controls
  - eMASS status update text
- **Footer:** CUI marking, page numbers

## Analysis Rules

### Severity Mapping
- High = CAT I
- Moderate = CAT II
- Low or Very Low = CAT III

### Risk Priority
- **Critical:** CAT I + 180+ days overdue
- **High:** CAT I + 30-179 days overdue
- **Medium:** CAT II + 90+ days overdue
- **Low:** All others

### Remediation Owner Assignment
Evaluated in order:
1. **ISO:** Contract, funding, procurement keywords
2. **ISSM:** IR-/CP- controls or incident/contingency keywords
3. **SysAdmin:** RA-5/SI-2 controls or scan/patch/STIG/SCAP keywords
4. **SysAdmin:** Hardware/switch/device/equipment keywords
5. **ISSO:** Documentation/policy/procedure keywords
6. **ISSO:** Default assignment

### Legacy/Unpatchable Detection
Flags systems containing:
- Windows XP, 7, 8, 8.1
- Windows Server 2003, 2008, 2012
- Specific build numbers (7600, 7601, 9600, etc.)
- EOL/end-of-life/unsupported keywords

## API Documentation

- [Ask Sage API Documentation](https://api.genai.army.mil/documentation/docs/api-documentation/api-documentation.html)

## License

GPL-2.0

## Author

Dan Barker

## Version

1.1
