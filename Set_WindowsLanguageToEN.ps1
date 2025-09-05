# requires -RunAsAdministrator
# Meysam 05-09-2025
# Log File in $Path1
# Versie: 1.0.5
# Table of Geographical location: https://learn.microsoft.com/en-us/windows/win32/intl/table-of-geographical-locations
<#
.SYNOPSIS
    Changes Windows display language and Welcome screen to en-US using .cab files from a GitHub repository or online sources.

.DESCRIPTION
    This script downloads .cab files from a public GitHub repository to C:\Temp\EN-US, installs the en-US language pack
    and additional Features on Demand (FoDs) offline, sets the display language for the current user, applies system-wide
    settings, and copies to the Welcome screen and new users. Falls back to online installation if needed. Requires a restart.

.PARAMETER LangCode
    The language code to install (default: en-US).

.PARAMETER GeoId
    The GeoId for the region (default: 244 for United States).

.PARAMETER RepositoryPath
    Local path for .cab files (default: C:\Temp\EN-US).

.PARAMETER GitHubRepoUrl
    Base URL of the GitHub repository containing .cab files (default: https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US).

.EXAMPLE
    # Run the script to download .cab files from GitHub and install en-US.
    .\Set-WindowsLanguage.ps1

.EXAMPLE
    # Run the script with local .cab files already in C:\Temp\EN-US.
    .\Set-WindowsLanguage.ps1 -RepositoryPath "C:\Temp\EN-US"
#>
param (
    [string]$LangCode = "en-US",
    [int]$GeoId = 244,  # GeoId for United States
    [string]$RepositoryPath = "C:\Temp\EN-US",  # Local download folder
    [string]$GitHubRepoUrl = "https://raw.githubusercontent.com/Meysam-Rajabipour/Language_Pack/main/LANG-Packages/EN-US"  # Base GitHub URL
)

# --- SCRIPT START ---
Write-Host "Starting Windows language configuration to $LangCode..." -ForegroundColor Cyan
$Path1 = "C:\Temp"
$logPath = Join-Path $Path1 "LanguageInstallLog.txt"
$xmlPath = Join-Path $env:TEMP "LanguageSettings.xml"
$registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$regKey = "DoNotConnectToWindowsUpdateInternetLocations"
$originalValue = $null

# Ensure log and download directories exist
if (-not (Test-Path $Path1)) {
    New-Item -Path $Path1 -ItemType Directory -Force
}
if (-not (Test-Path $RepositoryPath)) {
    New-Item -Path $RepositoryPath -ItemType Directory -Force
}

# Define .cab files to download and install
$cabFiles = @(
    "Microsoft-Windows-Client-Language-Pack_x64_$LangCode.cab",
    "Microsoft-Windows-LanguageFeatures-Basic-$LangCode-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-Handwriting-$LangCode-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-OCR-$LangCode-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-Speech-$LangCode-Package~31bf3856ad364e35~amd64~~.cab",
    "Microsoft-Windows-LanguageFeatures-TextToSpeech-$LangCode-Package~31bf3856ad364e35~amd64~~.cab"
)

$additionalCapabilityList = @(
    "Language.Basic~~~$LangCode~0.0.1.0",
    "Language.Handwriting~~~$LangCode~0.0.1.0",
    "Language.OCR~~~$LangCode~0.0.1.0",
    "Language.Speech~~~$LangCode~0.0.1.0",
    "Language.TextToSpeech~~~$LangCode~0.0.1.0"
)

try {
    #region 0. Download .cab Files from GitHub
    if ($GitHubRepoUrl) {
        Write-Host "`n[Step 0/6] Downloading .cab files from GitHub repository..." -ForegroundColor Yellow
        foreach ($cabFile in $cabFiles) {
            $url = "$GitHubRepoUrl/$cabFile"
            $outputPath = Join-Path $RepositoryPath $cabFile
            Write-Host "Downloading $cabFile from $url..."
            try {
                Invoke-WebRequest -Uri $url -OutFile $outputPath -ErrorAction Stop
                Write-Host "Downloaded $cabFile successfully." -ForegroundColor Green
            } catch {
                Write-Warning "Failed to download $cabFile from $url. Error: $_"
                if ($cabFile -eq "Microsoft-Windows-Client-Language-Pack_x64_$LangCode.cab") {
                    Write-Warning "Main language pack not found. Falling back to online installation..."
                    $GitHubRepoUrl = ""  # Disable GitHub download for main language pack
                    break
                } else {
                    Write-Warning "Skipping optional feature $cabFile."
                }
            }
        }
    }
    #endregion

    #region 1. Check Windows Update Connectivity (for online mode)
    if (-not $GitHubRepoUrl -and (-not $RepositoryPath -or -not (Test-Path $RepositoryPath))) {
        Write-Host "`n[Step 1/6] Checking Windows Update connectivity..." -ForegroundColor Yellow
        $wuTest = Test-NetConnection -ComputerName "www.update.microsoft.com" -Port 443
        if (-not $wuTest.TcpTestSucceeded) {
            Write-Error "Cannot connect to Windows Update (www.update.microsoft.com:443). Check internet or firewall settings."
            exit 1
        }
        Write-Host "Windows Update is accessible." -ForegroundColor Green
    }
    #endregion

    #region 2. Temporary WSUS Bypass (for online mode)
    if (-not $GitHubRepoUrl -and (-not $RepositoryPath -or -not (Test-Path $RepositoryPath))) {
        Write-Host "`n[Step 2/6] Temporarily disabling local update source restrictions..." -ForegroundColor Yellow
        $currentSetting = Get-ItemProperty -Path $registryPath -Name $regKey -ErrorAction SilentlyContinue
        if ($currentSetting) {
            $originalValue = $currentSetting.$regKey
            Set-ItemProperty -Path $registryPath -Name $regKey -Value 0
            Write-Host "Registry key '$regKey' set to 0." -ForegroundColor Green
        } else {
            Write-Host "No local update source restrictions found. No action needed." -ForegroundColor Green
        }
    }
    #endregion

    #region 3. Install Language Pack and FoDs
    Write-Host "`n[Step 3/6] Installing $LangCode language pack and additional features..." -ForegroundColor Yellow

    if ($RepositoryPath -and (Test-Path $RepositoryPath)) {
        # Offline installation using .cab files
        Write-Host "Using local repository at $RepositoryPath for offline installation..." -ForegroundColor Yellow

        # Install main language pack
        $langPackCab = Join-Path $RepositoryPath "Microsoft-Windows-Client-Language-Pack_x64_$LangCode.cab"
        if (Test-Path $langPackCab) {
            $dismCommand = "DISM /Online /Add-Package /PackagePath:`"$langPackCab`" /NoRestart /LogPath:`"$logPath`""
            Write-Host "Executing: $dismCommand"
            Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "DISM failed to install language pack with error code $LASTEXITCODE. Falling back to online installation..."
                $RepositoryPath = ""
            } else {
                Write-Host "$LangCode language pack installed successfully from local repository." -ForegroundColor Green
            }
        } else {
            Write-Warning "Language pack .cab file not found in $RepositoryPath. Falling back to online installation..."
            $RepositoryPath = ""
        }

        # Install additional FoDs
        foreach ($cabFile in $cabFiles | Where-Object { $_ -notlike "*Client-Language-Pack*" }) {
            $feature = Join-Path $RepositoryPath $cabFile
            if (Test-Path $feature) {
                $dismCommand = "DISM /Online /Add-Package /PackagePath:`"$feature`" /NoRestart /LogPath:`"$logPath`""
                Write-Host "Executing: $dismCommand"
                Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
                if ($LASTEXITCODE -ne 0) {
                    Write-Warning "Failed to install feature $feature with error code $LASTEXITCODE."
                } else {
                    Write-Host "Feature $feature installed successfully." -ForegroundColor Green
                }
            } else {
                Write-Warning "Feature file $feature not found in repository."
            }
        }

        # Install language group fonts (if applicable)
        $langGroup = ($LangCode -eq "en-US") ? "Latn" : ""  # en-US uses Latin fonts
        if ($langGroup) {
            $dismCommand = "DISM /Online /Add-Capability /CapabilityName:Language.Fonts.$langGroup~~~und-$langGroup~0.0.1.0 /Source:`"$RepositoryPath`" /NoRestart /LogPath:`"$logPath`""
            Write-Host "Executing: $dismCommand"
            Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Failed to install font group $langGroup with error code $LASTEXITCODE."
            } else {
                Write-Host "Font group $langGroup installed successfully." -ForegroundColor Green
            }
        }
    }

    if (-not $RepositoryPath -or -not (Test-Path $RepositoryPath)) {
        # Online installation
        Write-Host "Downloading and installing $LangCode language pack and features from Windows Update..." -ForegroundColor Yellow
        $dismCommand = "DISM /Online /Add-Capability /CapabilityName:Language.Basic~~~$LangCode~0.0.1.0 /Source:WindowsUpdate /NoRestart /LogPath:`"$logPath`""
        Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
        if ($LASTEXITCODE -ne 0) {
            Write-Warning "DISM online installation failed with error code $LASTEXITCODE. Attempting fallback method..."
            try {
                Install-Language -Language $LangCode
                $language = Get-WinUserLanguageList
                $language.Add($LangCode)
                Set-WinUserLanguageList -LanguageList $language -Force
            } catch {
                Write-Error "Fallback failed: Language pack not installed. Check $logPath. Aborting script."
                exit 1
            }
        }

        # Install additional capabilities
        foreach ($capability in $additionalCapabilityList) {
            $dismCommand = "DISM /Online /Add-Capability /CapabilityName:$capability /Source:WindowsUpdate /NoRestart /LogPath:`"$logPath`""
            Write-Host "Executing: $dismCommand"
            Invoke-Expression $dismCommand | Out-File -FilePath $logPath -Append
            if ($LASTEXITCODE -ne 0) {
                Write-Warning "Failed to install capability $capability with error code $LASTEXITCODE."
            } else {
                Write-Host "Capability $capability installed successfully." -ForegroundColor Green
            }
        }
    }
    #endregion

    #region 4. Verify Language Pack
    Write-Host "`n[Step 4/6] Verifying language pack installation..." -ForegroundColor Yellow
    $installedPackages = Get-WindowsPackage -Online | Where-Object { $_.PackageName -like "*$LangCode*" }
    if ($installedPackages -and $installedPackages.Count -gt 0) {
        Write-Host "Language pack for $LangCode verified." -ForegroundColor Green
    } else {
        Write-Error "Verification FAILED: Language pack for $LangCode not found. Check $logPath. Aborting script."
        exit 1
    }
    #endregion

    #region 5. Apply Language Settings
    Write-Host "`n[Step 5/6] Applying $LangCode settings for display, Welcome screen, and new users..." -ForegroundColor Yellow

    # Set display language for current user
    $langList = New-WinUserLanguageList -Language $LangCode
    Set-WinUserLanguageList $langList -Force
    Set-WinUILanguageOverride -Language $LangCode

    # Set system locale and home location
    Set-WinSystemLocale -SystemLocale $LangCode
    Set-WinHomeLocation -GeoId $GeoId

    # Configure Welcome screen and new user accounts via XML
    $xmlContent = @"
<gs:GlobalizationServices xmlns:gs="urn:longhornGlobalizationSvcs">
    <gs:UserList>
        <gs:User UserID="Current" CopySettingsToDefaultUserAcct="true" CopySettingsToSystemAcct="true"/>
    </gs:UserList>
    <gs:WindowsSettings>
        <gs:InputPreferences>
            <gs:InputPreference Language="$LangCode"/>
        </gs:InputPreferences>
        <gs:LocaleName Name="$LangCode"/>
    </gs:WindowsSettings>
</gs:GlobalizationSvcs>
"@
    $xmlContent | Out-File -FilePath $xmlPath -Encoding utf8
    Start-Process -FilePath "control.exe" -ArgumentList "intl.cpl,,/f:`"$xmlPath`"" -Wait

    # Additional copy to ensure Welcome screen and new users
    Copy-UserInternationalSettingsToSystem -WelcomeScreen $true -NewUser $true
    Write-Host "$LangCode settings applied for display, Welcome screen, and new user accounts." -ForegroundColor Green
    #endregion

    #region 6. Prompt for Restart
    Write-Host "`n[Step 6/6] Configuration complete. A restart is required to apply all changes." -ForegroundColor Cyan
    $choice = Read-Host "Do you want to restart the computer now? (Y/N)"
    if ($choice -match "^[Yy]") {
        Write-Host "Restarting computer..." -ForegroundColor Yellow
        Restart-Computer -Force
    } else {
        Write-Host "Restart cancelled. Please restart manually (or sign out/in) to apply all changes." -ForegroundColor Yellow
    }
    #endregion
}
catch {
    Write-Error "An error occurred: $_"
    Write-Host "Check $logPath for detailed logs." -ForegroundColor Yellow
    exit 1
}
finally {
    # Clean up XML file
    if (Test-Path $xmlPath) {
        Remove-Item $xmlPath -Force
        Write-Host "Cleaned up temporary XML file." -ForegroundColor Green
    }

    # Restore WSUS registry setting (if modified)
    if ($null -ne $originalValue) {
        Set-ItemProperty -Path $registryPath -Name $regKey -Value $originalValue
        Write-Host "Restored registry setting '$regKey' to original value." -ForegroundColor Green
    }
}

##EOF