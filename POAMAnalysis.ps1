<#
.SYNOPSIS
    POA&M Analysis Tool - Processes POA&M Excel files and generates ICS cybersecurity assessment reports

.DESCRIPTION
    This script extracts POA&M data from Excel files, analyzes each record using the Ask Sage GenAI API,
    and generates a Word document report with CUI markings.

.PREREQUISITES
    - Microsoft Excel installed
    - Microsoft Word installed
    - Ask Sage API key (see setup instructions below)
    - PowerShell 5.1 or higher
    - eMass Executive Report POA&Ms Export All POAM&s Excel file

.SETUP INSTRUCTIONS
    1. OBTAIN API KEY:
       - Navigate to Settings in Ask Sage https://api.genai.army.mil
       - Switch to the Account tab
       - Scroll down to "Manage your API Keys" in the sidebar
       - Click to generate your new API key
       - SECURITY: Keep your API key secure - treat it like a password and rotate regularly
    
    2. SET ENVIRONMENT VARIABLE:
       Run this command in PowerShell (replace YOUR_API_KEY with your actual key):
       
       [System.Environment]::SetEnvironmentVariable("GENAI_API_TOKEN", "YOUR_API_KEY", "User")
       
       Then restart PowerShell for the change to take effect.
    
    3. VERIFY SETUP:
       Test that the variable is set:
       
       [System.Environment]::GetEnvironmentVariable("GENAI_API_TOKEN", "User")

.USAGE
    1. Run the script: .\POAMAnalyzer.ps1
    2. Select your Excel POA&M file when prompted
    3. Choose output location for the Word document
    4. Wait for processing to complete
    5. Review the .docx output

.EXCEL FILE FORMAT
    - First two rows: Ignored (removed automatically)
    - Third row: Headers (must include: Organization, System Name, ID, Controls / APs, POA&M Item Status, 
      Scheduled Completion Date, Vulnerability Description, Severity, etc.)
    - Data rows: POA&M records

.API DOCUMENTATION
    https://api.genai.army.mil/documentation/docs/api-documentation/api-documentation.html

.NOTES
    Author: Dan Barker
    Liscence: GPL 2
    Version: 1.1
    Classification: Script is Unclass
    
.EXAMPLE
    .\POAMAnalyzer.ps1
    
    Runs the script interactively with file picker dialogs
#>

$now = Get-Date -Format 'yyyyMMdd-HHmm'

# ── Bitness guard ──────────────────────────────────────────────────────────────
# AF/DoD Microsoft 365 is nearly always 32-bit. When this script runs under
# 64-bit PowerShell the Excel COM type library can't load and throws:
#   TYPE_E_CANTLOADLIBRARY (0x80029C4A)
# Fix: detect 64-bit host and re-launch transparently under 32-bit PowerShell.
if ([System.IntPtr]::Size -eq 8) {
    $wow64PS = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path $wow64PS) {
        Write-Host "64-bit PowerShell detected — relaunching as 32-bit for Office COM compatibility..." -ForegroundColor Yellow
        Start-Process $wow64PS -ArgumentList "-File `"$PSCommandPath`"" -Wait
        exit
    }
}
# ──────────────────────────────────────────────────────────────────────────────

# Load Windows Forms assembly (ISE loads this automatically; console PowerShell does not)
Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# File picker for Excel input file
$OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$OpenFileDialog.Title = "Select Excel Input File"
$OpenFileDialog.Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*"
$OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")

# Create a hidden top-most form to own the dialogs (forces them to foreground in console PS)
$topForm = New-Object System.Windows.Forms.Form
$topForm.TopMost = $true

if ($OpenFileDialog.ShowDialog($topForm) -eq [System.Windows.Forms.DialogResult]::OK) {
    $ExcelFilePath = $OpenFileDialog.FileName
    Write-Host "Selected input file: $ExcelFilePath" -ForegroundColor Green
} else {
    Write-Host "No input file selected. Exiting." -ForegroundColor Yellow
    exit
}

# File picker for output Word file
$SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$SaveFileDialog.Title = "Select Output Word Document Location"
$SaveFileDialog.Filter = "Word Documents (*.docx)|*.docx|All Files (*.*)|*.*"
$SaveFileDialog.DefaultExt = "docx"
$SaveFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
$SaveFileDialog.FileName = "$now-POAMAnalysis.docx"

if ($SaveFileDialog.ShowDialog($topForm) -eq [System.Windows.Forms.DialogResult]::OK) {
    $OutputFilePath = $SaveFileDialog.FileName
    Write-Host "Selected output file: $OutputFilePath" -ForegroundColor Green
} else {
    Write-Host "No output file selected. Exiting." -ForegroundColor Yellow
    exit
}

$topForm.Dispose()

# Static prompt with current date embedded
$StaticPrompt = @"
You are an SME in ICS cybersecurity analyzing individual FRCS POA&M records.

=== REFERENCE DATE ===
Today's Date: $now
Use this date for overdue calculations

=== ANALYSIS RULES ===

1. SEVERITY MAPPING
   - High = CAT I
   - Moderate = CAT II  
   - Low or Very Low = CAT III
   - If Severity is blank, use Raw Severity
2. OVERDUE CALCULATION
   - Days Overdue = Today's Date minus Scheduled Completion Date
   - If Scheduled Completion Date is blank or future, Days Overdue = 0
3. CONTROL FAMILY EXTRACTION
   - Extract 2-letter prefix from Controls / APs (e.g., CM from CM-7.1)
   - List all unique control families in this POA&M
4. LEGACY/UNPATCHABLE DETECTION
   - Flag as legacy/unpatchable if Vulnerability Description OR Mitigations contains:
     Windows XP, Windows 7, Windows 8, Windows 8.1
     Windows Server 2003, Windows Server 2008, Windows Server 2012
     Build 7600, Build 7601, Build 9600, Build 10586, Build 17763, Build 18363, Build 19042
     EOL, end of life, end-of-life, unsupported, no longer supported
     legacy, no vendor support, obsolete
5. REMEDIATION OWNER ASSIGNMENT (evaluate in order)
   - If Comments/Recommendations contains contract, funding, procurement, CONS → ISO
   - If Controls / APs contains IR- or CP- OR Vulnerability Description contains incident or contingency → ISSM
   - If Controls / APs contains RA-5 or SI-2 OR Vulnerability Description contains scan, patch, STIG, SCAP → SysAdmin
   - If Vulnerability Description contains documentation, policy, procedure, plan → ISSO
   - If Vulnerability Description contains hardware, switch, device, equipment → SysAdmin
   - Default → ISSO
6. TEXT REPLACEMENT
   - Replace ALL instances of Thule with Pituffik
=== OUTPUT FORMAT ===
Plain text only - no tables, no code blocks, no markdown formatting
For the POA&M record provided, output ONLY the following:
---
POA&M ID: [13-digit ID]
POA&M URL:
Organization: [Base name]
System: [System Name] ([System Acronym])
Severity: [CAT I/II/III]
Status: [POA&M Item Status]
Scheduled Completion: [DD-MMM-YYYY format]
Days Overdue: [number or 0]
Control Families: [2-letter prefixes, comma-separated]
Legacy/Unpatchable: [Yes or No]
Risk Priority: [Select ONE: Critical, High, Medium, Low]
  - Critical = CAT I + 180+ days overdue
  - High = CAT I + 90+ days overdue OR CAT I + 30+ days overdue
  - Medium = CAT II + 180+ days overdue OR CAT II + 90+ days overdue
  - Low = All others
Vulnerability: [1-2 sentence summary from Vulnerability Description]
Recommended Fix: [1-2 sentences, ICS-safe approach. Prefer compensating controls, network segmentation, data diodes, application allowlisting. Reference AFCEC DET 5 FRCS Playbook or UFC 4-010-06 where applicable. Avoid automatic patching]
Remediation Owner: [ISO/ISSM/SysAdmin/ISSO per rules above]
Escalation Required: [Yes if CAT I and 90+ days overdue, otherwise No]
Compensating Control: [Cite specific control e.g. AC-4(21) via existing boundary protection, or N/A if full fix feasible]
eMASS Status Update:
[ICS-safe remediation action in present tense]. [Owner role] completing implementation with ETA [DD-MMM-YYYY, minimum 30 days from today]. [Compensating Control ID] enforced via [existing tool/mechanism]. Milestone [X]% complete. 
Next Review: [DD-MMM-YYYY, 30 days from today].
---

CONSTRAINTS:
- Do NOT invent data not in the record
- Use DD-MMM-YYYY date format
- If field is blank, state "Not specified"
- Plain text only

Analyze this POA&M record:
"@

# Define the essential fields we need (reduces payload size significantly)
$essentialFields = @(
    'Organization',
    'System Name',
    'ID',
    'POA&M URL',
    'Controls / APs',
    'POA&M Item Status',
    'Scheduled Completion Date',
    'Completion Date',
    'Risk Accepted Date',
    'Vulnerability Description',
    'Raw Severity',
    'Severity',
    'Mitigations',
    'Recommendations'
)

# Function to process a batch of rows
function Invoke-GenAIBatch {
    param(
        [array]$Rows,
        [int]$BatchNumber,
        [int]$TotalBatches
    )
    
    $headers = @{
        'x-access-tokens' = [System.Environment]::GetEnvironmentVariable("GENAI_API_TOKEN")
        'Content-Type' = 'application/json'
    }
    
    # Convert batch to JSON (already filtered to essential fields)
    $jsonData = $Rows | ConvertTo-Json -Depth 10 -Compress
    $fullPrompt = "$StaticPrompt`n`n$jsonData"
    
    Write-Host "  Prompt size: $($fullPrompt.Length) characters" -ForegroundColor Gray
    
    $body = @{
        message = $fullPrompt
        persona = 5
        model = 'google-claude-45-sonnet'
        conversation_type = 'CUI'
        temperature = 0.7
        limit_references = 5
        live = 0
    } | ConvertTo-Json -Depth 3
    
    try {
        Write-Host "Processing batch $BatchNumber of $TotalBatches..." -ForegroundColor Cyan
        $response = Invoke-RestMethod -Uri 'https://api.genai.army.mil/server/query' -Method Post -Headers $headers -Body $body -TimeoutSec 300
        Write-Host "  Batch $BatchNumber completed successfully" -ForegroundColor Green
        return $response.message
    }
    catch {
        Write-Error "  Batch $BatchNumber failed: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
            $responseBody = $reader.ReadToEnd()
            Write-Error "  Response: $responseBody"
        }
        return $null
    }
}

# Function to write content directly to Word document
function Add-ContentToWord {
    param(
        [object]$Selection,
        [string]$Content
    )
    
    # Split content into paragraphs
    $paragraphs = $Content -split "`n"
    
    foreach ($para in $paragraphs) {
        if ([string]::IsNullOrWhiteSpace($para)) {
            $Selection.TypeParagraph()
            continue
        }
        
        $trimmedPara = $para.Trim()
        
        # Detect POA&M record separators
        if ($trimmedPara -match '^---+$') {
            $Selection.TypeText("_" * 80)
            $Selection.TypeParagraph()
        }
        # Detect POA&M ID lines (for emphasis)
        elseif ($trimmedPara -match '^POA&M ID:') {
            $Selection.Font.Size = 12
            $Selection.Font.Bold = $true
            $Selection.Font.Color = 31  # Dark blue
            $Selection.TypeText($trimmedPara)
            $Selection.TypeParagraph()
            $Selection.Font.Size = 11
            $Selection.Font.Bold = $false
            $Selection.Font.Color = 0
        }
        # Detect field labels (text followed by colon)
        elseif ($trimmedPara -match '^([^:]+):\s*(.*)$') {
            $label = $Matches[1]
            $value = $Matches[2]
            
            # Color code based on priority/severity
            $labelColor = 0  # Default black
            if ($label -match 'Severity' -and $value -match 'CAT I') {
                $labelColor = 255  # Red for CAT I
            }
            elseif ($label -match 'Risk Priority' -and $value -match 'Critical') {
                $labelColor = 255  # Red for Critical
            }
            elseif ($label -match 'Risk Priority' -and $value -match 'High') {
                $labelColor = 0x0000FF  # Orange
            }
            
            $Selection.Font.Bold = $true
            $Selection.Font.Color = $labelColor
            $Selection.TypeText("$label`: ")
            $Selection.Font.Bold = $false
            $Selection.Font.Color = 0  # Black for value
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                $Selection.TypeText($value)
            }
            $Selection.TypeParagraph()
        }
        # Regular paragraph
        else {
            $Selection.TypeText($trimmedPara)
            $Selection.TypeParagraph()
        }
    }
}

# ── Open XML XLSX reader (no Excel COM required) ──────────────────────────────
# Reads .xlsx files directly as ZIP archives using built-in .NET types.
# Eliminates dependency on Excel COM automation and the TYPE_E_CANTLOADLIBRARY
# error caused by Office bitness mismatches or broken COM registrations.
function Read-XlsxData {
    param(
        [string]$FilePath,
        [string[]]$Fields
    )

    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $zip = [System.IO.Compression.ZipFile]::OpenRead($FilePath)
    try {
        # Shared string table (most cell text is stored here by reference)
        $ss = @()
        $ssEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/sharedStrings.xml' }
        if ($ssEntry) {
            $sr = New-Object System.IO.StreamReader($ssEntry.Open())
            [xml]$ssXml = $sr.ReadToEnd(); $sr.Close()
            $ss = @($ssXml.sst.si | ForEach-Object {
                if ($null -ne $_.t) {
                    if ($_.t -is [System.Xml.XmlElement]) { $_.t.InnerText } else { [string]$_.t }
                } elseif ($null -ne $_.r) {
                    (@($_.r) | ForEach-Object {
                        if ($_.t -is [System.Xml.XmlElement]) { $_.t.InnerText } else { [string]$_.t }
                    }) -join ''
                } else { '' }
            })
        }

        # Sheet 1 XML
        $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq 'xl/worksheets/sheet1.xml' }
        if (-not $sheetEntry) { throw "xl/worksheets/sheet1.xml not found — is this a valid .xlsx file?" }
        $sr = New-Object System.IO.StreamReader($sheetEntry.Open())
        [xml]$sx = $sr.ReadToEnd(); $sr.Close()

        # Column letters -> 1-based index (A=1, Z=26, AA=27 …)
        $colNumCache = @{}
        function Get-ColNum([string]$letters) {
            if ($colNumCache.ContainsKey($letters)) { return $colNumCache[$letters] }
            $n = 0
            foreach ($ch in $letters.ToUpper().ToCharArray()) { $n = $n * 26 + ([int][char]$ch - 64) }
            $colNumCache[$letters] = $n; $n
        }

        # Convert Excel date serial -> DD-MMM-YYYY when applicable
        $dateFields = 'Scheduled Completion Date','Completion Date','Risk Accepted Date'
        function Convert-ExcelDate([string]$v) {
            if ($v -match '^\d{4,5}$') {
                $serial = [int]$v
                if ($serial -gt 1 -and $serial -lt 55000) {
                    try { return ([datetime]'1899-12-30').AddDays($serial).ToString('dd-MMM-yyyy') } catch {}
                }
            }
            return $v
        }

        # Build a row-number -> { colIndex -> value } lookup
        $grid = @{}
        foreach ($row in $sx.worksheet.sheetData.row) {
            $ri = [int]$row.r
            $cells = @{}
            foreach ($c in $row.c) {
                if ($c.r -match '^([A-Z]+)\d+$') {
                    $ci = Get-ColNum $Matches[1]
                    $val = if ($c.t -eq 's' -and $null -ne $c.v) {
                        $ss[[int]$c.v]
                    } elseif ($c.t -eq 'inlineStr') {
                        if ($null -ne $c.is) { [string]$c.is.t } else { '' }
                    } elseif ($null -ne $c.v) { [string]$c.v } else { '' }
                    $cells[$ci] = [string]$val
                }
            }
            $grid[$ri] = $cells
        }

        # Row 3 is the header (rows 1-2 are eMass boilerplate)
        $headerRow = $grid[3]
        if (-not $headerRow) { throw "Header row not found at row 3. Check that this is an eMass Executive Report export." }

        $colToField = @{}
        foreach ($ci in $headerRow.Keys) {
            if ($Fields -contains $headerRow[$ci]) { $colToField[$ci] = $headerRow[$ci] }
        }

        Write-Host "Header mapping: found $($colToField.Count) of $($Fields.Count) essential fields" -ForegroundColor Green

        $maxRow = ($grid.Keys | Measure-Object -Maximum).Maximum
        $results = [System.Collections.Generic.List[hashtable]]::new()

        for ($r = 4; $r -le $maxRow; $r++) {
            $cells = $grid[$r]
            if (-not $cells) { continue }

            $obj = @{}
            $hasData = $false
            foreach ($f in $Fields) { $obj[$f] = '' }

            foreach ($ci in $colToField.Keys) {
                $fname = $colToField[$ci]
                $val = if ($cells.ContainsKey($ci)) { $cells[$ci] } else { '' }
                if ($dateFields -contains $fname) { $val = Convert-ExcelDate $val }
                $obj[$fname] = $val
                if ($val -ne '') { $hasData = $true }
            }

            if ($hasData) { $results.Add($obj) }

            if ($r % 10 -eq 0) {
                Write-Host "  Processed $($r - 3) of $($maxRow - 3) rows..." -ForegroundColor Gray
            }
        }

        return $results
    }
    finally {
        $zip.Dispose()
    }
}
# ──────────────────────────────────────────────────────────────────────────────

# Load and process Excel file
$word = $null
$doc = $null
$selection = $null

try {
    Write-Host "`nLoading Excel file (Open XML — no Office COM required)..." -ForegroundColor Cyan

    if ($ExcelFilePath -notmatch '\.xlsx$') {
        throw "Only .xlsx files are supported. Please save your eMass export as .xlsx and retry."
    }

    $allRows = Read-XlsxData -FilePath $ExcelFilePath -Fields $essentialFields

    Write-Host "Extracted $($allRows.Count) data rows (essential fields only)" -ForegroundColor Green
    
    # ====================================
    # INITIALIZE WORD DOCUMENT
    # ====================================
    
    Write-Host "`nInitializing Word document..." -ForegroundColor Cyan
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $selection = $word.Selection
    Write-Host "Word document created" -ForegroundColor Green
    
    # ====================================
    # DOCUMENT HEADER
    # ====================================
    
    Write-Host "Adding document header..." -ForegroundColor Cyan
    
    # Add classification marking at top
    $selection.Font.Size = 12
    $selection.Font.Bold = $true
    $selection.Font.Color = 255  # Red
    $selection.ParagraphFormat.Alignment = 1  # Center
    $selection.TypeText("CUI")
    $selection.TypeParagraph()
    $selection.TypeParagraph()
    
    # Add title
    $selection.Font.Size = 16
    $selection.Font.Bold = $true
    $selection.Font.Color = 0  # Black
    $selection.TypeText("FRCS POA&M Analysis Report")
    $selection.TypeParagraph()
    
    # Add date
    $selection.Font.Size = 11
    $selection.Font.Bold = $false
    $selection.TypeText("Generated: $now")
    $selection.TypeParagraph()
    $selection.TypeParagraph()
    
    # Reset formatting for body
    $selection.Font.Size = 11
    $selection.Font.Bold = $false
    $selection.Font.Color = 0  # Black
    $selection.ParagraphFormat.Alignment = 0  # Left
    
    Write-Host "Document header complete" -ForegroundColor Green
    
    # ====================================
    # PROCESS BATCHES AND WRITE TO WORD
    # ====================================
    
    $batchSize = 1 #This is all that works for now
    $totalBatches = [Math]::Ceiling($allRows.Count / $batchSize)
    
    Write-Host "`nProcessing $($allRows.Count) rows in $totalBatches batches (batch size: $batchSize)..." -ForegroundColor Yellow
    
    # Process in batches
    $successfulBatches = 0
    for ($i = 0; $i -lt $allRows.Count; $i += $batchSize) {
        $batchNumber = [Math]::Floor($i / $batchSize) + 1
        $endIndex = [Math]::Min($i + $batchSize - 1, $allRows.Count - 1)
        $batch = $allRows[$i..$endIndex]
        
        Write-Host "`n--- Batch $batchNumber of $totalBatches (rows $($i + 1) to $($endIndex + 1)) ---" -ForegroundColor Yellow
        
        $result = Invoke-GenAIBatch -Rows $batch -BatchNumber $batchNumber -TotalBatches $totalBatches
        
        if ($result) {
            $successfulBatches++
            Write-Host "  Writing to Word document..." -ForegroundColor Cyan
            Add-ContentToWord -Selection $selection -Content $result
            $selection.TypeParagraph()
            Write-Host "  Content added to Word document" -ForegroundColor Green
        } else {
            Write-Host "  Batch $batchNumber failed - continuing with next batch..." -ForegroundColor Red
        }
        
        # Rate limiting - wait between batches
        if ($batchNumber -lt $totalBatches) {
            Write-Host "  Waiting 3 seconds before next batch..." -ForegroundColor Gray
            Start-Sleep -Seconds 3
        }
    }
    
    # ====================================
    # DOCUMENT FOOTER
    # ====================================
    
    Write-Host "`nAdding document footer..." -ForegroundColor Cyan
    
    # Add page numbers
    $word.ActiveWindow.ActivePane.View.SeekView = 4  # wdSeekCurrentPageFooter
    $selection.ParagraphFormat.Alignment = 1  # Center
    
    # Add classification marking in footer
    $selection.Font.Size = 10
    $selection.Font.Bold = $true
    $selection.Font.Color = 255  # Red
    $selection.TypeText("CUI")
    $selection.TypeParagraph()
    
    # Add page numbers
    $selection.Font.Size = 10
    $selection.Font.Bold = $false
    $selection.Font.Color = 0  # Black
    $selection.TypeText("Page ")
    $selection.Fields.Add($selection.Range, 33) | Out-Null  # wdFieldPage
    $selection.TypeText(" of ")
    $selection.Fields.Add($selection.Range, 34) | Out-Null  # wdFieldNumPages
    
    $word.ActiveWindow.ActivePane.View.SeekView = 0  # wdSeekMainDocument
    
    Write-Host "Document footer complete" -ForegroundColor Green
    
    # ====================================
    # SAVE DOCUMENT
    # ====================================
    
    Write-Host "`nSaving Word document..." -ForegroundColor Cyan
    $doc.SaveAs([ref]$OutputFilePath, [ref]16)
    Write-Host "Word document saved successfully" -ForegroundColor Green
    
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "All Processing Complete!" -ForegroundColor Green
    Write-Host "Successful batches: $successfulBatches of $totalBatches" -ForegroundColor Green
    Write-Host "Word document: $OutputFilePath" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
}
catch {
    Write-Error "Error processing: $($_.Exception.Message)"
    Write-Error $_.ScriptStackTrace
}
finally {
    # Clean up COM objects
    Write-Host "`nCleaning up..." -ForegroundColor Cyan
    
    if ($doc) {
        $doc.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) | Out-Null
    }
    if ($word) {
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
    if ($selection) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($selection) | Out-Null
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
    Write-Host "Complete!" -ForegroundColor Green
}