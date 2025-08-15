<#
This script detects password-protected MS Office files, ZIP archives, and encrypted PDF files
in a SharePoint 2016 document library by inspecting binary header,
using 7-Zip for ZIP analysis, and static inspection for PDF encryption markers.

Detection Logic:
- Office 97–2003 formats (.doc, .xls, .ppt) are scanned using binary header offsets
- Office 2007+ formats (.docx, .xlsx, .pptx) are scanned for "EncryptedPackage" markers
- ZIP archives are scanned using 7-Zip command-line output for encryption flags
- PDF files are scanned for static encryption markers like /Encrypt and /Filter /Standard

Security:
- This script performs static analysis only
- It does NOT execute files or use COM automation (e.g., New-Object -ComObject Word.Application)
- This eliminates the risk of triggering warmful embedded code

Requirements:
- Run from a domain-joined workstation HRDC-DRHC
- CSOM DLLs must be present in .\CSOM-DLLs\
- 7z.exe must be available in system path or will be installed automatically from 7-Zip.org

Output:
- A timestamped CSV report listing scanned files and their encryption status
- A detailed log file capturing scan results and operational events

Supported Formats:
- Word 97–2003 (.doc, .dot), Word 2007+ (.docx, .docm, .dotm)
- Excel 97–2003 (.xls), Excel 2007+ (.xlsx, .xlsm, .xlsb)
- PowerPoint 97–2003 (.ppt), PowerPoint 2007+ (.pptx, .pptm, .ppsm)
- ZIP archives (.zip)
- PDF files (.pdf)

#>

# Discover script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Load SharePoint CSOM Assemblies from relative folder
Add-Type -Path (Join-Path $scriptDir "CSOM-DLLs\Microsoft.SharePoint.Client.dll")
Add-Type -Path (Join-Path $scriptDir "CSOM-DLLs\Microsoft.SharePoint.Client.Runtime.dll")

# Set variables
$siteUrl     = "<YOUR_SHAREPOINT_SITE_URL>"
$libraryName = "<YOUR_DOCUMENT_LIBRARY_NAME>"
$domain      = "<YOUR_DOMAIN>"
$tempPath    = Join-Path $env:TEMP "SharePointPWDprotectedFiles"
$logFile     = Join-Path $tempPath "PasswordProtectedFiles-Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
$csvFile     = Join-Path $tempPath "PasswordProtectedFiles-Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
$7zipExe     = "${env:ProgramFiles}\7-Zip\7z.exe"
$downloadPageUrl = "https://www.7-zip.org/download.html"
$installerPath = Join-Path $env:TEMP "7zip_installer.exe"

# Ensure temp folder exists
if (-not (Test-Path $tempPath)) {
    New-Item -Path $tempPath -ItemType Directory | Out-Null
}

# Logging function
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp - $message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# Function to get the installed version of 7-Zip
function Get-7ZipVersion {
    if (Test-Path $7zipExe) {
        $versionOutput = & $7zipExe | Select-String "7-Zip"
        if ($versionOutput) {
            return ($versionOutput -split " ")[2]
        }
    }
    return $null
}

# Check if 7-Zip is installed
$installedVersion = Get-7ZipVersion
if ($installedVersion) {
    Write-Host "7-Zip is installed. Version: $installedVersion"
    Write-Log "7-Zip detected. Version: $installedVersion"
} else {
    Write-Host "7-Zip is not installed. Installing now..."
    Write-Log "7-Zip not found. Attempting installation..."

    try {
        $downloadPageContent = Invoke-WebRequest -Uri $downloadPageUrl -UseBasicParsing
        $downloadUrl = $downloadPageContent.Links | Where-Object { $_.href -match "a/7z\d+-x64.exe" } | Select-Object -ExpandProperty href -First 1
        $downloadUrl = "https://www.7-zip.org/$downloadUrl"

        Write-Host "Downloading 7-Zip from $downloadUrl..."
        Invoke-WebRequest -Uri $downloadUrl -OutFile $installerPath

        Write-Host "Installing 7-Zip silently..."
        Start-Process -FilePath $installerPath -ArgumentList "/S" -Wait

        $newVersion = Get-7ZipVersion
        if ($newVersion) {
            Write-Host "7-Zip installed successfully. Version: $newVersion"
            Write-Log "7-Zip installed. Version: $newVersion"
        } else {
            Write-Host "Installation may have failed. 7-Zip version not detected."
            Write-Log "7-Zip installation failed or version not detected."
        }

        Remove-Item -Path $installerPath -Force
    } catch {
        Write-Host "Failed to install 7-Zip: $_"
        Write-Log "7-Zip installation error: $_"
    }
}

# Auth block using Windows Integrated Authentication
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$ctx.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials

$list = $ctx.Web.Lists.GetByTitle($libraryName)
Write-Log "Starting password protection scan in: $siteUrl/$libraryName"

# CAML Query
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

# Write CSV header
"Filename,Extension,SharePointURL,IsPwdProtected" | Out-File -FilePath $csvFile -Encoding UTF8

# Function to detect password protection
function find-docpasswords($ctx, $item) {
    $fileName  = $item["FileLeafRef"]
    $fileRef   = $item["FileRef"]
    $extension = [System.IO.Path]::GetExtension($fileName).ToLower()
    $url       = "$domain$($item.File.ServerRelativeUrl)"

    $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($ctx, $fileRef)
    $stream = New-Object System.IO.MemoryStream
    $fileInfo.Stream.CopyTo($stream)

    $bytes = $stream.ToArray()
    $isEncrypted = $false

    # Office 2007+ detection
    try {
        $startText = [System.Text.Encoding]::Default.GetString($bytes[0..2000]).Replace("`0", " ")
        if ($startText -like "*E n c r y p t e d P a c k a g e*") {
            $isEncrypted = $true
        }
    } catch {}

    # Legacy Office binary format detection
if (-not $isEncrypted -and $bytes.Length -gt 0x220) {
    $prefix = [System.Text.Encoding]::Default.GetString($bytes[0..1])
    if ($prefix -eq "ÐÏ") {
        # XLS 2003
        if ($bytes[0x208] -eq 0xFE) {
            $isEncrypted = $true
            Write-Log "$fileName → XLS 2003 password protection detected"
        }
        # XLS 2005
        elseif ($bytes[0x214] -eq 0x2F) {
            $isEncrypted = $true
            Write-Log "$fileName → XLS 2005 password protection detected"
        }
        # DOC 2005
        elseif ($bytes[0x20B] -eq 0x13) {
            $isEncrypted = $true
            Write-Log "$fileName → DOC 2005 password protection detected"
        }
        # Office 2007+ fallback
        elseif ($bytes.Length -gt 2000) {
            $startText = [System.Text.Encoding]::Default.GetString($bytes[0..2000]).Replace("`0", " ")
            if ($startText -like "*E n c r y p t e d P a c k a g e*") {
                $isEncrypted = $true
                Write-Log "$fileName → Office 2007+ encrypted package detected"
            }
        }
    }
}

    # ZIP file password protection detection using 7z.exe
    if (-not $isEncrypted -and $extension -eq ".zip") {
        try {
            $tempZipPath = Join-Path $tempPath ([System.Guid]::NewGuid().ToString() + "_" + $fileName)
            [System.IO.File]::WriteAllBytes($tempZipPath, $bytes)

            $listOutput = & $7zipExe l -slt "`"$tempZipPath`"" 2>&1
            if ($listOutput -match "Encrypted = \+") {
                $isEncrypted = $true
                Write-Log "$fileName → ZIP file is password protected"
            }

            Remove-Item $tempZipPath -Force
        } catch {
            Write-Log "ZIP scan failed for ${fileName}: $_"
        }
    }

    # PDF encryption detection using static markers
    if (-not $isEncrypted -and $extension -eq ".pdf") {
        try {
            $pdfText = [System.Text.Encoding]::ASCII.GetString($bytes)
            if ($pdfText -match "/Encrypt" -and ($pdfText -match "/AuthEvent\s*/DocOpen" -or $pdfText -match "/Filter\s*/Standard")) {
                $isEncrypted = $true
                Write-Log "$fileName PDF file is encrypted"
            }
        } catch {
            Write-Log "PDF scan failed for ${fileName}: $_"
        }
    }

    $status = if ($isEncrypted) { "Yes" } else { "No" }

    # Write to CSV
    "$fileName,$extension,$url,$status" | Out-File -FilePath $csvFile -Append -Encoding UTF8

    # Log result
    Write-Log "$fileName → Password Protected: $status"

    $stream.Close()
    $fileInfo.Dispose()
}

# Loop through items and scan each document
foreach ($item in $items) {
    $file = $item.File
    $ctx.Load($file)
    $ctx.ExecuteQuery()

    find-docpasswords $ctx $item
}

Write-Log "Password protection scan completed. CSV report saved to: $csvFile"