# requires -RunAsAdministrator
# Meysam 05-09-2025
# Log File in C:\Temp
# Versie: 1.0.9V004
<#
.SYNOPSIS
    Downloads en-US language pack and FoD .cab files from GitHub based on ListCabfiles.txt, skipping existing files.

.DESCRIPTION
    Downloads ListCabfiles.txt from GitHub, reads .cab file names, downloads only files not already in C:\Temp\EN-US,
    installs and verifies packages, and applies en-US language settings. Logs to C:\Temp\LanguageDownloadLog.txt.

.PARAMETER RepositoryPath
    Local path to save and use .cab files (default: C:\Temp\EN-US).

.PARAMETER GitHubRepoUrl
    Base URL of the GitHub repository (default: https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US).

.EXAMPLE
    .\DownloadAndApply-LanguageCabs.ps1
#>
param (
    [string]$RepositoryPath = "C:\Temp\EN-US",
    [string]$GitHubRepoUrl = "https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US",
    [string]$LangCode = "en-US"
)

# --- SCRIPT START ---
Write-Host "Starting .cab file download and language pack installation for $LangCode..." -ForegroundColor Cyan
$Path1 = "C:\Temp"
$logPath = Join-Path $Path1 "LanguageDownloadLog.txt"
$listFilePath = Join-Path $Path1 "ListCabfiles.txt"
##TEST
# Ensure log and download directories exist
if (-not (Test-Path $Path1)) {
    New-Item -Path $Path1 -ItemType Directory -Force
}
if (-not (Test-Path $RepositoryPath)) {
    New-Item -Path $RepositoryPath -ItemType Directory -Force
}

try {
    # Download ListCabfiles.txt
    Write-Host "`n[Step 1/5] Downloading ListCabfiles.txt from $GitHubRepoUrl..." -ForegroundColor Yellow
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
        $cabFiles = Get-Content -Path $listFilePath | Where-Object { $_ -like "*.cab" -or $_ -like "*.Appx" } | ForEach-Object { $_.Trim() }
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

    # Download .cab files listed in ListCabfiles.txt, skipping existing files
    Write-Host "`n[Step 2/5] Downloading .cab files from GitHub repository..." -ForegroundColor Yellow
    foreach ($cabFile in $cabFiles) {
        $outputPath = Join-Path $RepositoryPath $cabFile
        if (Test-Path $outputPath) {
            Write-Host "$cabFile already exists in $RepositoryPath. Skipping download." -ForegroundColor Green
            "$cabFile already exists in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
            continue
        }

        $url = "$GitHubRepoUrl/$cabFile"
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
    Write-Host "`n[Step 3/5] Verifying downloaded files..." -ForegroundColor Yellow
    $downloadedFiles = Get-ChildItem -Path $RepositoryPath -Filter "*.cab"
    if ($downloadedFiles.Count -gt 0) {
        Write-Host "Downloaded files:" -ForegroundColor Green
        $downloadedFiles | ForEach-Object { Write-Host $_.Name }
        "Verified $($downloadedFiles.Count) .cab files in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    } else {
        Write-Warning "No .cab files were downloaded to $RepositoryPath."
        "No .cab files found in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }

    # Install .cab files and check installation status
    Write-Host "`n[Step 4/5] Installing .cab files from $RepositoryPath..." -ForegroundColor Yellow
    foreach ($cabFile in $cabFiles) {
        $filePath = Join-Path $RepositoryPath $cabFile
        $packageName = [System.IO.Path]::GetFileNameWithoutExtension($cabFile)

        # Check if package is already installed
        $isInstalled = Get-WindowsPackage -Online | Where-Object { $_.PackageName -like "*$packageName*" -and $_.PackageState -eq "Installed" }
        if ($isInstalled) {
            Write-Host "$cabFile is already installed. Skipping." -ForegroundColor Green
            "Package $cabFile is already installed at $(Get-Date)" | Out-File -FilePath $logPath -Append
            continue
        }

        if (Test-Path $filePath) {
            $dismCommand = "DISM /Online /Add-Package /PackagePath:`"$filePath`" /NoRestart /LogPath:`"$logPath`""
            Write-Host "Executing: $dismCommand"
            $output = Invoke-Expression $dismCommand 2>&1
            $output | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Failed to install $cabFile with error code $LASTEXITCODE."
                if ($LASTEXITCODE -eq -2146498536) {
                    Write-Warning "Error 0x80240028: $cabFile is not applicable to this Windows version."
                }
                "Failed to install $cabFile with error code $LASTEXITCODE at $(Get-Date)" | Out-File -FilePath $logPath -Append
            } else {
                # Verify installation
                $isInstalled = Get-WindowsPackage -Online | Where-Object { $_.PackageName -like "*$packageName*" -and $_.PackageState -eq "Installed" }
                if ($isInstalled) {
                    Write-Host "$cabFile installed and verified successfully." -ForegroundColor Green
                    "Package $cabFile installed and verified at $(Get-Date)" | Out-File -FilePath $logPath -Append
                } else {
                    Write-Warning "$cabFile installation completed but verification failed."
                    "Package $cabFile installation completed but not found in installed packages at $(Get-Date)" | Out-File -FilePath $logPath -Append
                }
            }
        } else {
            Write-Warning "$cabFile not found in $RepositoryPath."
            "$cabFile not found in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
        }
    }

    # Install font group
    Write-Host "`nInstalling font group for $LangCode..." -ForegroundColor Yellow
    $langGroup = if ($LangCode -eq "en-US") { "Latn" } else { "" }
    if ($langGroup) {
        $fontCapability = "Language.Fonts.$langGroup~~~und-$langGroup~0.0.1.0"
        $isFontInstalled = Get-WindowsCapability -Online | Where-Object { $_.Name -eq $fontCapability -and $_.State -eq "Installed" }
        if ($isFontInstalled) {
            Write-Host "Font group $langGroup is already installed. Skipping." -ForegroundColor Green
            "Font group $langGroup is already installed at $(Get-Date)" | Out-File -FilePath $logPath -Append
        } else {
            $dismCommand = "DISM /Online /Add-Capability /CapabilityName:$fontCapability /Source:`"$RepositoryPath`" /NoRestart /LogPath:`"$logPath`""
            Write-Host "Executing: $dismCommand"
            $output = Invoke-Expression $dismCommand 2>&1
            $output | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Failed to install font group $langGroup with error code $LASTEXITCODE."
                "Failed to install font group $langGroup with error code $LASTEXITCODE at $(Get-Date)" | Out-File -FilePath $logPath -Append
            } else {
                $isFontInstalled = Get-WindowsCapability -Online | Where-Object { $_.Name -eq $fontCapability -and $_.State -eq "Installed" }
                if ($isFontInstalled) {
                    Write-Host "Font group $langGroup installed and verified successfully." -ForegroundColor Green
                    "Font group $langGroup installed and verified at $(Get-Date)" | Out-File -FilePath $logPath -Append
                } else {
                    Write-Warning "Font group $langGroup installation completed but verification failed."
                    "Font group $langGroup installation completed but not found in installed capabilities at $(Get-Date)" | Out-File -FilePath $logPath -Append
                }
            }
        }
    }

    # Apply language settings
    Write-Host "`n[Step 5/5] Applying $LangCode language settings..." -ForegroundColor Yellow
    $langList = New-WinUserLanguageList -Language $LangCode
    Set-WinUserLanguageList $langList 
    Set-WinUILanguageOverride -Language $LangCode
    Set-WinSystemLocale -SystemLocale $LangCode

    # Prompt for restart
    Write-Host "`nConfiguration complete. Restart required." -ForegroundColor Cyan
    $choice = Read-Host "Restart now? (Y/N)"
    if ($choice -match "^[Yy]") {
        Write-Host "Restarting computer..." -ForegroundColor Yellow
        #Restart-Computer -Force
    } else {
        #Write-Host "Restart cancelled. Please restart manually." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Error: $_"
    Write-Host "Check $logPath for details." -ForegroundColor Yellow
    exit 1
}


Install-language -language $LangCode
#Install-languagefeatures -language $LangCode
Write-Host "`nLanguage pack installation process completed. Check $logPath for details." -ForegroundColor Cyan
##EOF