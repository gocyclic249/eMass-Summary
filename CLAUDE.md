# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**eMass-Summary** is a single PowerShell script (`POAMAnalysis.ps1`) that reads POA&M (Plan of Actions and Milestones) data from eMass Executive Report Excel files, sends each record to the Ask Sage GenAI API (Claude 4.5 Sonnet), and generates a formatted, CUI-marked Word document with cybersecurity assessment analysis.

> The script itself is UNCLASSIFIED. It becomes CUI once eMass data is loaded.

## Running the Script

```powershell
.\POAMAnalysis.ps1
```

No build step. The script launches Windows file-picker dialogs to select the input Excel file and output `.docx` path.

**One-time API key setup:**
```powershell
[System.Environment]::SetEnvironmentVariable("GENAI_API_TOKEN", "YOUR_API_KEY", "User")
# Restart PowerShell, then verify:
[System.Environment]::GetEnvironmentVariable("GENAI_API_TOKEN", "User")
```

## Linting

```powershell
Invoke-ScriptAnalyzer -Path .\POAMAnalysis.ps1
```

Requires `PSScriptAnalyzer` (`Install-Module PSScriptAnalyzer`). No test suite exists.

## Architecture

The script has three sequential phases:

1. **Data Extraction** — Opens the Excel file via COM object, strips header rows, and extracts 14 essential fields per POA&M record into a PowerShell object array.

2. **AI Analysis** — Iterates records one at a time (`$batchSize = 1`), sending each to `https://api.genai.army.mil/server/query` via `Invoke-RestMethod`. A static system prompt (lines ~107–179) encodes all analysis rules. A 3-second delay is enforced between requests.

3. **Document Output** — Parses the AI response text and writes it to a Word COM object with CUI header/footer markings and severity-based color coding (CAT I → red, High priority → orange).

### Key Functions

| Function | Location | Purpose |
|---|---|---|
| `Invoke-GenAIBatch` | ~line 200 | Builds the API request payload and POSTs to Ask Sage |
| `Add-ContentToWord` | ~line 245 | Parses the AI response and writes formatted paragraphs to the Word document |

### Embedded Analysis Rules (in the static prompt)

- **Severity mapping**: High → CAT I, Moderate → CAT II, Low/Very Low → CAT III
- **Days overdue**: `Today - Scheduled Completion Date` (0 if future or blank)
- **Risk priority**: Critical (CAT I + 180+ days), High (CAT I + 30–179 days), Medium (CAT II + 90+ days), Low (all others)
- **Remediation owner**: Evaluated in priority order — ISO (contract/funding), ISSM (IR-/CP- controls), SysAdmin (RA-5/SI-2, patch/STIG/scan), ISSO (documentation/policy, default fallback)
- **Legacy detection**: Flags EOL OS strings (Windows XP/7/8/Server 2003/2008/2012, etc.)

### API Configuration (hardcoded)

```powershell
$apiUrl          = "https://api.genai.army.mil/server/query"
$model           = "google-claude-45-sonnet"
$conversationType = "CUI"
$temperature     = 0.7
$batchSize       = 1
```

## Platform Constraints

- **Windows only** — depends on Excel and Word COM objects (`New-Object -ComObject`)
- Requires Microsoft Office installed locally
- PowerShell 5.1+ (note: interactive file dialogs work in PowerShell console but historically required ISE — see commit `33b676a` for the fix)
