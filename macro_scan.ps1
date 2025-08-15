<#
This script scans Microsoft Office files stored in a SharePoint 2016 document library for embedded macros.
It invokes a Python-based scanner (OleMacroDetector.py) to perform static analysis
using multiple techniques including olevba, oledump, and VBA_Parser.

Detection Methods:
- olevba: triage scan for macro indicators (autoexec, VBA)
- oledump.py: checks for macro streams in OLE containers
- VBA_Parser: deep inspection of macro content across formats (OLE, OpenXML, MHTML)
- Excel 97 OLE stream check: verifies presence of macro-related streams
- ZIP-based scan: inspects embedded OLE files inside zipped Office formats

OleMacroDetector.py:
This Python script acts as a macro triage engine. It combines multiple static analysis tools to detect macros
in Office documents and returns a simple result:
- "MACROS_PRESENT" if any macro indicators are found
- "NO_MACROS" otherwise

Security Note
This script performs **static analysis only**. It does not execute macro code or use COM automation (e.g., `New-Object -ComObject Word.Application`),
which eliminates the risk of triggering malicious macros during scanning.

Supported Formats:
- Word 97–2003 (.doc, .dot), Word 2007+ (.docm, .dotm)
- Excel 97–2003 (.xls), Excel 2007+ (.xlsm, .xlsb)
- PowerPoint 97–2003 (.ppt), PowerPoint 2007+ (.pptm, .ppsm)
- Word/PowerPoint 2007+ XML (Flat OPC)
- Word 2003 XML (.xml)
- Word/Excel Single File Web Page / MHTML (.mht)
- Publisher (.pub)
- SYLK/SLK files (.slk)
- Text files containing VBA or VBScript source code
- Password-protected ZIP archives containing any of the above

Requirements:
- Run from a domain-joined workstation HRDC-DRHC
- SharePoint 2016 CSOM DLLs must be present in .\CSOM-DLLs\
- Python must be installed and available in system PATH
- Required Python packages: olefile, oletools
- oledump.py will be downloaded automatically from GitHub if not already present
- OleMacroDetector.py must be present in the same folder as this PowerShell script
#>

# Discover script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load SharePoint CSOM Assemblies from relative folder
Add-Type -Path (Join-Path $scriptDir "CSOM-DLLs\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Join-Path $scriptDir "CSOM-DLLs\Microsoft.SharePoint.Client.Runtime.dll")

# Set variables
$siteUrl     = "<YOUR_SHAREPOINT_SITE_URL>"
$libraryName = "<YOUR_DOCUMENT_LIBRARY_NAME>"
$timestamp   = Get-Date -Format 'yyyyMMdd_HHmmss'
$tempPath    = Join-Path $env:TEMP "SharePointOfficeMacroScan"

# Output file paths stored in temp folder
$logFile     = Join-Path $tempPath "MacroScan_Office_$timestamp.txt"
$csvFile     = Join-Path $tempPath "MacroScan_Office_$timestamp.csv"

# Python script location (same folder as PowerShell script)
$pythonScript = Join-Path $scriptDir "OleMacroDetector.py"

# Supported macro-aware extensions
$macroExtensions = ".xls", ".doc", ".ppt", ".xlsm", ".xlsx", ".docm", ".dotm", ".pptm", ".potm", ".ppsm", ".sldm"

function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

try {
    $pythonPath = (Get-Command python -ErrorAction Stop).Path
    Write-Host "Python found at: $pythonPath"

    $version = (&{python -V} 2>&1)
    Write-Host "Python version: $version"
} catch {
    Write-Host "Python not found in PATH. Please install Python to proceed."
    exit 1
}

# Location to save oledump.py
$oledumpPath = Join-Path $tempPath "oledump.py"
$oledumpUrl = "https://github.com/DidierStevens/DidierStevensSuite/raw/refs/heads/master/oledump.py"

if (-Not (Test-Path $oledumpPath)) {
    try {
        Write-Host "Downloading oledump.py from GitHub..."
        Invoke-WebRequest -Uri $oledumpUrl -OutFile $oledumpPath -UseBasicParsing
        Write-Host "Download completed: $oledumpPath"
    } catch {
        Write-Host "Failed to download oledump.py: $($_.Exception.Message)"
        Write-Log "Download error: $($_.Exception.Message)"
    }
} else {
    Write-Host "oledump.py already exists at: $oledumpPath"
}

# Auth block
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$ctx.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

$list = $ctx.Web.Lists.GetByTitle($libraryName)
Write-Log "Starting macro scan in: $siteUrl/$libraryName"

$camlQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
$camlQuery.ViewXml = @"
<View Scope='RecursiveAll'>
    <Query><OrderBy><FieldRef Name='ID' Ascending='TRUE'/></OrderBy></Query>
    <RowLimit Paged='TRUE'>5000</RowLimit>
</View>
"@
$items = $list.GetItems($camlQuery)
$ctx.Load($items)
$ctx.ExecuteQuery()

New-Item -Path $tempPath -ItemType Directory -Force | Out-Null
"Filename,Extension,SharePointURL,HasMacro" | Out-File -FilePath $csvFile -Encoding UTF8

function Check-Macros {
    param ($localFile)

    try {
        $result = & python $pythonScript $localFile 2>&1
        Write-Log "Python scan result for ${localFile}: $result"
        if ($result -match "MACROS_PRESENT") {
            return "Yes"
        }
    } catch {
        Write-Log "Error scanning ${localFile}: $($_.Exception.Message)"
    }

    return "No"
}

foreach ($item in $items) {
    $fileRef   = $item["FileRef"]
    $fileName  = $item["FileLeafRef"]
    $createdBy = $item["Author"].Email
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()

    if ($macroExtensions -contains $extension) {
        # Load the File object separately to avoid "pending query" error
        $file = $item.File
        $ctx.Load($file)
        $ctx.ExecuteQuery()  
        
        $filePath = Join-Path $tempPath $fileName
        $stream   = New-Object System.IO.MemoryStream
        $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($ctx, $fileRef)
        $fileInfo.Stream.CopyTo($stream)
        [System.IO.File]::WriteAllBytes($filePath, $stream.ToArray())
        $stream.Close()

        $hasMacro = Check-Macros -localFile $filePath

        # Extract full URL using domain + server-relative path
        $domain = "<YOUR_DOMAIN>"
        $url = "$domain$($item.File.ServerRelativeUrl)"

        "$fileName,$extension,$url,$hasMacro,$fileSizeBytes,$createdBy" | Out-File -FilePath $csvFile -Append -Encoding UTF8
        Write-Log "Scan complete: $fileName → Macro: $hasMacro"

        Remove-Item $filePath -Force
    }
}

Write-Log "Macro scan completed. CSV report saved to: $csvFile"