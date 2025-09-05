# requires -RunAsAdministrator
# Meysam 05-09-2025
# Log File in $Path1
# Versie: 1.0.9
<#
.SYNOPSIS
    Downloads en-US language pack and Feature on Demand (FoD) .cab files from a GitHub repository to C:\Temp\EN-US.

.DESCRIPTION
    This script downloads specified .cab files from a public GitHub repository to the local folder C:\Temp\EN-US.
    It logs all download attempts and errors to C:\Temp\LanguageDownloadLog.txt.

.PARAMETER RepositoryPath
    Local path to save downloaded .cab files (default: C:\Temp\EN-US).

.PARAMETER GitHubRepoUrl
    Base URL of the GitHub repository containing .cab files (default: https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US).

.EXAMPLE
    # Download .cab files from the default GitHub repository.
    .\Download-LanguageCabs.ps1

.EXAMPLE
    # Download .cab files to a custom folder.
    .\Download-LanguageCabs.ps1 -RepositoryPath "C:\CustomFolder"
#>
param (
    [string]$RepositoryPath = "C:\Temp\EN-US",  # Local download folder
    [string]$GitHubRepoUrl = "https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US"  # Base GitHub URL
)

# --- SCRIPT START ---
Write-Host "Starting .cab file download for en-US language pack..." -ForegroundColor Cyan
$Path1 = "C:\Temp"
$logPath = Join-Path $Path1 "LanguageDownloadLog.txt"

# Ensure log and download directories exist
if (-not (Test-Path $Path1)) {
    New-Item -Path $Path1 -ItemType Directory -Force
}
if (-not (Test-Path $RepositoryPath)) {
    New-Item -Path $RepositoryPath -ItemType Directory -Force
}

# Define .cab files to download
try {
    # Download ListCabfiles.txt
    Write-Host "`nDownloading ListCabfiles.txt from $GitHubRepoUrl..." -ForegroundColor Yellow
    $listUrl = "$GitHubRepoUrl/ListCabfiles.txt"
    try {
        Invoke-WebRequest -Uri $listUrl -OutFile $listFilePath -ErrorAction Stop
        Write-Host "Downloaded ListCabfiles.txt successfully to $listFilePath." -ForegroundColor Green
        "Downloaded ListCabfiles.txt successfully to $listFilePath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    } catch {
        Write-Error "Failed to download ListCabfiles.txt from $listUrl. Error: $_"
        "Failed to download ListCabfiles.txt from $listUrl. Error: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
        exit 1
    }

    # Read .cab file names from ListCabfiles.txt
    if (Test-Path $listFilePath) {
        $cabFiles = Get-Content -Path $listFilePath | Where-Object { $_ -like "*.cab" } | ForEach-Object { $_.Trim() }
        if (-not $cabFiles) {
            Write-Error "No .cab files listed in ListCabfiles.txt or file is empty."
            "No .cab files listed in ListCabfiles.txt or file is empty at $(Get-Date)" | Out-File -FilePath $logPath -Append
            exit 1
        }
        Write-Host "Found $($cabFiles.Count) .cab files in ListCabfiles.txt:" -ForegroundColor Green
        $cabFiles | ForEach-Object { Write-Host $_ }
    } else {
        Write-Error "ListCabfiles.txt not found at $listFilePath."
        "ListCabfiles.txt not found at $listFilePath at $(Get-Date)" | Out-File -FilePath $logPath -Append
        exit 1
    }

    # Download .cab files listed in ListCabfiles.txt
    Write-Host "`nDownloading .cab files from GitHub repository..." -ForegroundColor Yellow
    foreach ($cabFile in $cabFiles) {
        $url = "$GitHubRepoUrl/$cabFile"
        $outputPath = Join-Path $RepositoryPath $cabFile
        Write-Host "Downloading $cabFile from $url..."
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
            Write-Host "Downloaded $cabFile successfully to $outputPath." -ForegroundColor Green
            "Downloaded $cabFile successfully to $outputPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
        } catch {
            Write-Warning "Failed to download $cabFile from $url. Error: $_"
            "Failed to download $cabFile from $url. Error: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
        }
    }

    # Verify downloaded files
    Write-Host "`nVerifying downloaded files..." -ForegroundColor Yellow
    $downloadedFiles = Get-ChildItem -Path $RepositoryPath -Filter "*.cab"
    if ($downloadedFiles.Count -gt 0) {
        Write-Host "`nDownloaded files:" -ForegroundColor Green
        $downloadedFiles | ForEach-Object { Write-Host $_.Name }
        "Verified $($downloadedFiles.Count) .cab files in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    } else {
        Write-Warning "No .cab files were downloaded to $RepositoryPath."
        "No .cab files found in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }
}
catch {
    Write-Error "An error occurred during download: $_"
    "Error during download: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
    exit 1
}

Write-Host "`nDownload process completed. Check $logPath for details." -ForegroundColor Cyan

###APPLY LANGUAGE PACK USING DISM###
$LangCode = "en-US"

# --- SCRIPT START ---
Write-Host "Applying $LangCode language pack..." -ForegroundColor Cyan
$logPath = "C:\Temp\LanguageInstallLog.txt"

# Ensure log directory exists
$Path1 = "C:\Temp"
if (-not (Test-Path $Path1)) {
    New-Item -Path $Path1 -ItemType Directory -Force
}

# Define .cab files
$cabFiles = @(
    "Microsoft-Windows-Client-Language-Pack_x64_en-us.cab",
    "Microsoft-Windows-LanguageFeatures-Basic-en-us-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-Handwriting-en-us-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-OCR-en-us-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-Speech-en-us-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-TextToSpeech-en-us-Package~31bf3856ad364e35~amd64~~.cab"
)


try {
    # Install .cab files
    Write-Host "Installing .cab files from $RepositoryPath..." -ForegroundColor Yellow
    foreach ($cabFile in $cabFiles) {
        $filePath = Join-Path $RepositoryPath $cabFile
        if (Test-Path $filePath) {
            $dismCommand = "DISM /Online /Add-Package /PackagePath:`"$filePath`"  /LogPath:`"$logPath`""
            Write-Host "`nExecuting: $dismCommand"
            Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "`nFailed to install $cabFile with error code $LASTEXITCODE."
            } else {
                Write-Host "$cabFile `ninstalled successfully." -ForegroundColor Green
            }
        } else {
            Write-Warning "$cabFile not found in $RepositoryPath."
        }
    }

    # Install font group
    $langGroup = if ($LangCode -eq "en-US") { "Latn" } else { "" }
    if ($langGroup) {
        $dismCommand = "DISM /Online /Add-Capability /CapabilityName:Language.Fonts.$langGroup~~~und-$langGroup~0.0.1.0 /Source:`"$RepositoryPath`" /NoRestart /LogPath:`"$logPath`""
        Write-Host "`nExecuting: $dismCommand"
        Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "Failed to install font group $langGroup with error code $LASTEXITCODE."
        } else {
            Write-Host "Font group $langGroup installed successfully." -ForegroundColor Green
        }
    }

    # Apply language settings
    Write-Host "Applying $LangCode language settings..." -ForegroundColor Yellow
    $langList = New-WinUserLanguageList -Language $LangCode
    Set-WinUserLanguageList $langList -Force
    Set-WinUILanguageOverride -Language $LangCode
    Set-WinSystemLocale -SystemLocale $LangCode

    # Prompt for restart
    Write-Host "`nConfiguration complete. Restart required." -ForegroundColor Cyan
    $choice = Read-Host "Restart now? (Y/N)"
    if ($choice -match "^[Yy]") {
        Write-Host "Restarting computer..." -ForegroundColor Yellow
        Restart-Computer -Force
    } else {
        Write-Host "Restart cancelled. Please restart manually." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Error: $_"
    Write-Host "Check $logPath for details." -ForegroundColor Yellow
    exit 1
}
Write-Host "`nLanguage pack installation process completed. Check $logPath for details." -ForegroundColor Cyan