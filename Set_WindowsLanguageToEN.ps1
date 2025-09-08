# requires -RunAsAdministrator
# Meysam 05-09-2025
# Log File in C:\Temp
# Versie: 1.0.9V005
<#
.SYNOPSIS
    Downloads and applies en-US language pack, FoD .cab files, and LanguageExperiencePack .Appx from GitHub.

.DESCRIPTION
    Pulls latest repository changes, downloads ListCabfiles.txt, downloads missing .cab and .Appx files to C:\Temp\EN-US,
    installs and verifies packages, and applies en-US language settings. Logs to C:\Temp\LanguageDownloadLog.txt.

.PARAMETER RepositoryPath
    Local path to save and use files (default: C:\Temp\EN-US).

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
Write-Host "Starting file download and language pack installation for $LangCode..." -ForegroundColor Cyan
$Path1 = "C:\Temp"
$logPath = Join-Path $Path1 "LanguageDownloadLog.txt"
$listFilePath = Join-Path $Path1 "ListCabfiles.txt"
$gitRepoPath = "C:\temp\git.repo"

# Ensure log and download directories exist
if (-not (Test-Path $Path1)) {
    New-Item -Path $Path1 -ItemType Directory -Force
}
if (-not (Test-Path $RepositoryPath)) {
    New-Item -Path $RepositoryPath -ItemType Directory -Force
}

try {
    # Update local Git repository to ensure latest ListCabfiles.txt
    Write-Host "`n[Step 1/6] Updating local Git repository at $gitRepoPath..." -ForegroundColor Yellow
    if (Test-Path $gitRepoPath) {
        Push-Location $gitRepoPath
        try {
            git pull origin main 2>&1 | Out-File -FilePath $logPath -Append
            Write-Host "Git repository updated successfully." -ForegroundColor Green
            "Git repository updated successfully at $(Get-Date)" | Out-File -FilePath $logPath -Append
        } catch {
            Write-Warning "Failed to update Git repository. Proceeding with existing files. Error: $_"
            "Failed to update Git repository. Error: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
        }
        Pop-Location
    } else {
        Write-Warning "Git repository not found at $gitRepoPath. Skipping Git update."
        "Git repository not found at $gitRepoPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }

    # Download ListCabfiles.txt
    Write-Host "`n[Step 2/6] Downloading ListCabfiles.txt from $GitHubRepoUrl..." -ForegroundColor Yellow
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

    # Read .cab and .Appx file names from ListCabfiles.txt
    if (Test-Path $listFilePath) {
        $packageFiles = Get-Content -Path $listFilePath | Where-Object { $_ -like "*.cab" -or $_ -like "*.Appx" } | ForEach-Object { $_.Trim() }
        if (-not $packageFiles) {
            Write-Error "No .cab or .Appx files listed in ListCabfiles.txt or file is empty."
            "No .cab or .Appx files listed in ListCabfiles.txt or file is empty at $(Get-Date)" | Out-File -FilePath $logPath -Append
            exit 1
        }
        Write-Host "Found $($packageFiles.Count) .cab and .Appx files in ListCabfiles.txt:" -ForegroundColor Green
        $packageFiles | ForEach-Object { Write-Host $_ }
    } else {
        Write-Error "ListCabfiles.txt not found at $listFilePath."
        "ListCabfiles.txt not found at $listFilePath at $(Get-Date)" | Out-File -FilePath $logPath -Append
        exit 1
    }

    # Download .cab and .Appx files, skipping existing ones
    Write-Host "`n[Step 3/6] Downloading .cab and .Appx files from GitHub repository..." -ForegroundColor Yellow
    foreach ($packageFile in $packageFiles) {
        $outputPath = Join-Path $RepositoryPath $packageFile
        if (Test-Path $outputPath) {
            Write-Host "$packageFile already exists in $RepositoryPath. Skipping download." -ForegroundColor Green
            "$packageFile already exists in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
            continue
        }

        $url = "$GitHubRepoUrl/$packageFile"
        Write-Host "Downloading $packageFile from $url..."
        try {
            Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
            Write-Host "Downloaded $packageFile successfully to $outputPath." -ForegroundColor Green
            "Downloaded $packageFile successfully to $outputPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
        } catch {
            Write-Warning "Failed to download $packageFile from $url. Error: $_"
            "Failed to download $packageFile from $url. Error: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
        }
    }

    # Verify downloaded files
    Write-Host "`n[Step 4/6] Verifying downloaded files..." -ForegroundColor Yellow
    $downloadedFiles = Get-ChildItem -Path $RepositoryPath -Filter "*.*" | Where-Object { $_.Extension -in ".cab", ".Appx" }
    if ($downloadedFiles.Count -gt 0) {
        Write-Host "Downloaded files:" -ForegroundColor Green
        $downloadedFiles | ForEach-Object { Write-Host $_.Name }
        "Verified $($downloadedFiles.Count) .cab and .Appx files in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    } else {
        Write-Warning "No .cab or .Appx files were downloaded to $RepositoryPath."
        "No .cab or .Appx files found in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }

    # Install .cab and .Appx files and check installation status
    Write-Host "`n[Step 5/6] Installing .cab and .Appx files from $RepositoryPath..." -ForegroundColor Yellow
    foreach ($packageFile in $packageFiles) {
        $filePath = Join-Path $RepositoryPath $packageFile
        $packageName = [System.IO.Path]::GetFileNameWithoutExtension($packageFile)

        if (-not (Test-Path $filePath)) {
            Write-Warning "$packageFile not found in $RepositoryPath."
            "$packageFile not found in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
            continue
        }

        if ($packageFile -like "*.cab") {
            # Check if .cab package is already installed
            $isInstalled = Get-WindowsPackage -Online | Where-Object { $_.PackageName -like "*$packageName*" -and $_.PackageState -eq "Installed" }
            if ($isInstalled) {
                Write-Host "$packageFile is already installed. Skipping." -ForegroundColor Green
                "Package $packageFile is already installed at $(Get-Date)" | Out-File -FilePath $logPath -Append
                continue
            }

            $dismCommand = "DISM /Online /Add-Package /PackagePath:`"$filePath`" /NoRestart /LogPath:`"$logPath`""
            Write-Host "Executing: $dismCommand"
            $output = Invoke-Expression $dismCommand 2>&1
            $output | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Failed to install $packageFile with error code $LASTEXITCODE."
                if ($LASTEXITCODE -eq -2146498536) {
                    Write-Warning "Error 0x80240028: $packageFile is not applicable to this Windows version."
                }
                "Failed to install $packageFile with error code $LASTEXITCODE at $(Get-Date)" | Out-File -FilePath $logPath -Append
            } else {
                # Verify .cab installation
                $isInstalled = Get-WindowsPackage -Online | Where-Object { $_.PackageName -like "*$packageName*" -and $_.PackageState -eq "Installed" }
                if ($isInstalled) {
                    Write-Host "$packageFile installed and verified successfully." -ForegroundColor Green
                    "Package $packageFile installed and verified at $(Get-Date)" | Out-File -FilePath $logPath -Append
                } else {
                    Write-Warning "$packageFile installation completed but verification failed."
                    "Package $packageFile installation completed but not found in installed packages at $(Get-Date)" | Out-File -FilePath $logPath -Append
                }
            }
        } elseif ($packageFile -like "*.Appx") {
            # Check if .Appx package is already installed
            $isAppxInstalled = Get-AppxPackage -Name "Microsoft.LanguageExperiencePacken-US" | Where-Object { $_.PackageFullName -like "*$packageName*" }
            if ($isAppxInstalled) {
                Write-Host "$packageFile is already installed. Skipping." -ForegroundColor Green
                "Package $packageFile is already installed at $(Get-Date)" | Out-File -FilePath $logPath -Append
                continue
            }

            Write-Host "Installing $packageFile with Add-AppxPackage..."
            try {
                Add-AppxPackage -Path $filePath -ErrorAction Stop
                Write-Host "$packageFile installed successfully." -ForegroundColor Green
                "Package $packageFile installed successfully at $(Get-Date)" | Out-File -FilePath $logPath -Append
            } catch {
                Write-Warning "Failed to install $packageFile. Error: $_"
                "Failed to install $packageFile. Error: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
            }
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
    Write-Host "`n[Step 6/6] Applying $LangCode language settings..." -ForegroundColor Yellow
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
##EOF