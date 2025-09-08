# requires -RunAsAdministrator
# Meysam 05-09-2025
# Log File in C:\Temp
# Versie: 1.0.9V006
<#
.SYNOPSIS
    Downloads and applies en-US language pack, FoD .cab files, and LanguageExperiencePack .Appx from GitHub.

.DESCRIPTION
    Pulls latest repository changes, downloads ListCabfiles.txt, downloads missing .cab and .Appx files,
    installs packages, enables side-loading, installs optional features, applies administrative language
    settings via XML, and verifies installations. Logs to C:\Temp\LanguageDownloadLog.txt.

.PARAMETER RepositoryPath
    Local path to save and use files (default: C:\Temp\EN-US).

.PARAMETER GitHubRepoUrl
    Base URL of the GitHub repository (default: https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US).

.PARAMETER LangCode
    Language code to apply (default: en-US).

.PARAMETER PrimaryInputCode
    Input language ID (default: 0409:00000409 for en-US).

.PARAMETER PrimaryGeoID
    GeoID for location (default: 176 for en-US).

.EXAMPLE
    .\DownloadAndApply-LanguageCabs.ps1
#>
param (
    [string]$RepositoryPath = "C:\Temp\EN-US",
    [string]$GitHubRepoUrl = "https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US",
    [string]$LangCode = "en-US",
    [string]$PrimaryInputCode = "0409:00000409",
    [string]$PrimaryGeoID = "176"
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
    # Enable side-loading for .Appx files
    Write-Host "`n[Step 1/7] Enabling side-loading for .Appx files..." -ForegroundColor Yellow
    New-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock -Name AllowAllTrustedApps -Value 1 -PropertyType DWORD -Force | Out-Null
    Write-Host "Side-loading enabled." -ForegroundColor Green
    "Side-loading enabled at $(Get-Date)" | Out-File -FilePath $logPath -Append

    # Update local Git repository
     Write-Host "`n[Step 2/7] Updating repository at $gitRepoPath..." -ForegroundColor Yellow
    
    # Download ListCabfiles.txt
    Write-Host "`n[Step 3/7] Downloading ListCabfiles.txt from $GitHubRepoUrl..." -ForegroundColor Yellow
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
    Write-Host "`n[Step 4/7] Downloading .cab and .Appx files from GitHub repository..." -ForegroundColor Yellow
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
    Write-Host "`n[Step 5/7] Verifying downloaded files..." -ForegroundColor Yellow
    $downloadedFiles = Get-ChildItem -Path $RepositoryPath -Filter "*.*" | Where-Object { $_.Extension -in ".cab", ".Appx" }
    if ($downloadedFiles.Count -gt 0) {
        Write-Host "Downloaded files:" -ForegroundColor Green
        $downloadedFiles | ForEach-Object { Write-Host $_.Name }
        "Verified $($downloadedFiles.Count) .cab and .Appx files in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    } else {
        Write-Warning "No .cab or .Appx files were downloaded to $RepositoryPath."
        "No .cab or .Appx files found in $RepositoryPath at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }

    # Install .cab and .Appx files
    Write-Host "`n[Step 6/7] Installing .cab and .Appx files from $RepositoryPath..." -ForegroundColor Yellow
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

    # Install optional features for the primary language
    Write-Host "`nInstalling optional features for $LangCode..." -ForegroundColor Yellow
    $capabilities = Get-WindowsCapability -Online | Where-Object { $_.Name -like "*$LangCode*" -and $_.State -ne "Installed" }
    if ($capabilities) {
        foreach ($capability in $capabilities) {
            Write-Host "Installing capability $($capability.Name)..."
            try {
                Add-WindowsCapability -Online -Name $capability.Name -ErrorAction Stop
                Write-Host "Capability $($capability.Name) installed successfully." -ForegroundColor Green
                "Capability $($capability.Name) installed successfully at $(Get-Date)" | Out-File -FilePath $logPath -Append
            } catch {
                Write-Warning "Failed to install capability $($capability.Name). Error: $_"
                "Failed to install capability $($capability.Name). Error: $_ at $(Get-Date)" | Out-File -FilePath $logPath -Append
            }
        }
    } else {
        Write-Host "No additional optional features to install for $LangCode." -ForegroundColor Green
        "No additional optional features to install for $LangCode at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }

    # Apply custom XML for administrative language settings
    Write-Host "`n[Step 7/7] Applying administrative language settings via XML..." -ForegroundColor Yellow
    $xmlPath = Join-Path $env:TEMP "en-US.xml"
    $XML = @"
<gs:GlobalizationServices xmlns:gs="urn:longhornGlobalizationUnattend">
    <gs:UserList>
        <gs:User UserID="Current" CopySettingsToDefaultUserAcct="true" CopySettingsToSystemAcct="true"/> 
    </gs:UserList>
    <gs:LocationPreferences> 
        <gs:GeoID Value="$PrimaryGeoID"/>
    </gs:LocationPreferences>
    <gs:MUILanguagePreferences>
        <gs:MUILanguage Value="$LangCode"/>
    </gs:MUILanguagePreferences>
    <gs:SystemLocale Name="$LangCode"/>
    <gs:InputPreferences>
        <gs:InputLanguageID Action="add" ID="$PrimaryInputCode" Default="true"/>
    </gs:InputPreferences>
    <gs:UserLocale>
        <gs:Locale Name="$LangCode" SetAsCurrent="true" ResetAllSettings="false"/>
    </gs:UserLocale>
</gs:GlobalizationServices>
"@
    New-Item -Path $xmlPath -ItemType File -Value $XML -Force | Out-Null
    Write-Host "Created XML file at $xmlPath."
    "Created XML file at $xmlPath at $(Get-Date)" | Out-File -FilePath $logPath -Append

    $process = Start-Process -FilePath Control.exe -ArgumentList "intl.cpl,,/f:`"$xmlPath`"" -NoNewWindow -PassThru -Wait
    if ($process.ExitCode -eq 0) {
        Write-Host "Administrative language settings applied successfully." -ForegroundColor Green
        "Administrative language settings applied successfully at $(Get-Date)" | Out-File -FilePath $logPath -Append
    } else {
        Write-Warning "Failed to apply administrative language settings. Exit code: $($process.ExitCode)"
        "Failed to apply administrative language settings. Exit code: $($process.ExitCode) at $(Get-Date)" | Out-File -FilePath $logPath -Append
    }
    Remove-Item -Path $xmlPath -Force -ErrorAction SilentlyContinue

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