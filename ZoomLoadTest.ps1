<#
    Author      : Benjamin TAN
    Date        : 26 Feb 2025
    Purpose     : Automate Zoom installation, stress test, and cleanup with logging and CPU/memory failsafe.
    Version     : 1.0.0
    CreatedDate : 2025-02-25
    LastUpdated : 2025-02-27 22:14:00
#>


# Get current script directory
$CurrentPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$DownloadURL = "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64"
$InstallerPath = "$CurrentPath\ZoomInstaller.msi"
$ZoomPath = "C:\Program Files\Zoom\bin\zoom.exe"  # Update if necessary
$ZoomDir = "C:\Program Files\Zoom\bin"  # Folder containing zoom executables
$LogFile = "$CurrentPath\ZoomAutomation.log"
$ConfigFile = "$CurrentPath\ZoomConfig.json"
$LastUsedNumberFile = "$CurrentPath\LastUsedNumber.txt"

# Function to write logs with timestamps
function Write-Log {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp - $Message"
    Add-Content -Path $LogFile -Value $LogEntry
    Write-Output $LogEntry
}

# Function to read config file
function Read-Config {
    if (Test-Path $ConfigFile) {
        return Get-Content $ConfigFile | ConvertFrom-Json
    }
    return $null
}

# Function to save config file
function Save-Config {
    param ([string]$MeetingID, [string]$MeetingCode)
    $ConfigData = @{
        MeetingID   = $MeetingID
        MeetingCode = $MeetingCode
    }
    $ConfigData | ConvertTo-Json | Set-Content $ConfigFile
    Write-Log "Saved Meeting ID ($MeetingID) and Code ($MeetingCode) to config file."
}

function Get-NextInstanceNumber {
    if (Test-Path $LastUsedNumberFile) {
        $LastUsedNumber = Get-Content $LastUsedNumberFile
        return [int]$LastUsedNumber + 1
    }
    return 1  # If file doesn't exist, start with 1
}

function Update-LastUsedNumber {
    param ([int]$Number)
    Set-Content -Path $LastUsedNumberFile -Value $Number
}

# Function to disable Windows Firewall for all profiles
function Disable-Firewall {
    Write-Output "Disabling Windows Firewall..."

    # Disable firewall for all profiles
    Set-NetFirewallProfile -Profile Domain,Private,Public -Enabled False

    Write-Output "Windows Firewall has been disabled."
    Write-Log "Windows Firewall has been disabled."
}

# Function to enable Windows Firewall for all profiles
function Enable-Firewall {
    Write-Output "Enabling Windows Firewall..."

    # Enable firewall for all profiles
    Set-NetFirewallProfile -Profile Domain,Private,Public -Enabled True

    Write-Output "Windows Firewall has been enabled."
    Write-Log "Windows Firewall has been enabled."

}

# Function to stop all running Zoom processes
function Stop-AllZoomProcesses {
    Write-Log "Stopping all Zoom processes..."
    $ZoomProcesses = Get-Process | Where-Object { $_.ProcessName -match "^zoom(\d+)?$" }
    
    if ($ZoomProcesses) {
        $ZoomProcesses | ForEach-Object { Stop-Process -Id $_.Id -Force }
        Write-Log "All Zoom processes have been stopped."
    } else {
        Write-Log "No Zoom processes were running."
    }
}

# Function to start Zoom stress load with CPU and memory failsafe
function Start-ZoomStressLoad {
    if (-Not (Test-Path $ZoomPath)) {
        Write-Log "ERROR: Zoom.exe not found at $ZoomPath. Please install Zoom first."
        return
    }

    # Check for saved meeting details
    $Config = Read-Config
    if ($Config -and $Config.MeetingID -and $Config.MeetingCode) {
        $UseConfig = Read-Host "Use saved meeting details? (Y/N) Meeting ID: $($Config.MeetingID), Code: $($Config.MeetingCode)"
    
        # Check if the user input is valid
        if ($UseConfig -match "^[Yy]$") {
            $MeetingID = $Config.MeetingID
            $MeetingCode = $Config.MeetingCode
            Write-Log "Using saved meeting details."
        }
        elseif ($UseConfig -match "^[Nn]$") {
            Write-Log "User chose not to use saved configuration. Proceeding with manual input."
        }
        else {
            Write-Log "Invalid input detected. The configuration will not be overwritten."
            $MeetingID = $Config.MeetingID
            $MeetingCode = $Config.MeetingCode
        }
    } else {
        Write-Log "No saved meeting configuration found. Proceeding with manual input."
    }

    # Prompt user for new meeting details if needed
    if (-not $MeetingID -or -not $MeetingCode) {
        $MeetingID = Read-Host "Enter the Meeting ID"
        $MeetingCode = Read-Host "Enter the Meeting Code"
        Save-Config -MeetingID $MeetingID -MeetingCode $MeetingCode
        Write-Log "Saved Meeting ID ($MeetingID) and Code ($MeetingCode) to config file."
    }

    # Get the next instance number
    $NextInstanceNumber = Get-NextInstanceNumber
    Write-Log "Next available instance number: $NextInstanceNumber"

    $RunCount = Read-Host "Enter the number of times to copy and run Zoom.exe"
    if ($RunCount -match '^\d+$') {
        $RunCount = [int]$RunCount  
        Write-Log "Starting Zoom stress test with $RunCount instances (Meeting ID: $MeetingID, Code: $MeetingCode)."

        for ($i = 0; $i -lt $RunCount; $i++) {
            $InstanceNumber = ($NextInstanceNumber + $i).ToString("D2")

            # Check CPU and memory usage before launching a new instance
            $CPU_Usage = (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue
            $Memory_Usage = (Get-Counter '\Memory\% Committed Bytes In Use').CounterSamples.CookedValue
            
            if ($CPU_Usage -gt 90 -or $Memory_Usage -gt 90) {
                Write-Log "WARNING: High resource usage detected! CPU: $CPU_Usage% | Memory: $Memory_Usage%"
                $UserChoice = Read-Host "System resources are over 90%. Continue launching Zoom instances? (Y/N)"
                
                if ($UserChoice -match "^[Nn]$") {
                    Write-Log "User chose to stop Zoom stress test due to high CPU/memory usage."
                    return
                }
            }

            # Copy zoom.exe for stress testing
            $NewZoomPath = "$ZoomDir\zoom$InstanceNumber.exe"
            Copy-Item -Path $ZoomPath -Destination $NewZoomPath -Force
            Write-Log "Copied zoom.exe to zoom$InstanceNumber.exe"

            # Start the copied Zoom instance with the meeting link
            $MeetingURL = "zoommtg://zoom.us/join?action=join&uname=User$InstanceNumber&confno=$MeetingID&pwd=$MeetingCode&opt=join&role=0"
            Start-Process -FilePath $NewZoomPath -ArgumentList "--url=`"$MeetingURL`"" -NoNewWindow
            Write-Log "Started zoom$InstanceNumber.exe with meeting link: $MeetingURL"
        }

        # Update the last used number after starting the instances
        $LastUsedNumber = $NextInstanceNumber + $RunCount - 1
        Update-LastUsedNumber -Number $LastUsedNumber
    } else {
        Write-Log "ERROR: Invalid input for number of instances."
    }
}

# Function to stop Zoom stress load and delete copied Zoom files
function Stop-ZoomStressLoad {
    Write-Log "Stopping all Zoom stress test instances..."
    Stop-AllZoomProcesses

    Write-Log "Awaiting the termination of a hooked process."
    for ($i = 5; $i -gt 0; $i--) { Write-Host "$i..."; Start-Sleep -Seconds 1 }

    # Delete copied Zoom files
    Write-Log "Deleting copied zoom.exe files..."
    $CopiedZoomFiles = Get-ChildItem -Path $ZoomDir -Filter "zoom*.exe" | Where-Object { $_.Name -match "^zoom\d+\.exe$" }

    if ($CopiedZoomFiles) {
        $CopiedZoomFiles | ForEach-Object {
            Remove-Item -Path $_.FullName -Force
            Write-Log "Deleted: $($_.FullName)"
        }
        Write-Log "All copied Zoom files have been deleted."
    } else {
        Write-Log "No copied Zoom files found."
    }

    # Reset the last used instance number to 0
    Write-Log "Resetting last used instance number to 0."
    Update-LastUsedNumber -Number 0
}

# Function to download Zoom installer
function Download-Zoom {
    Write-Log "Starting Zoom download to $InstallerPath..."
    $ProgressPreference = 'SilentlyContinue'

    try {
        Invoke-WebRequest -Uri $DownloadURL -OutFile $InstallerPath
        Write-Log "Zoom downloaded successfully to $InstallerPath"
    } catch {
        Write-Log "ERROR: Failed to download Zoom - $_"
    }
}

# Function to install Zoom silently
function Install-Zoom {
    Stop-AllZoomProcesses  # Ensure Zoom is not running before installing

    if (-Not (Test-Path $InstallerPath)) {
        Write-Log "Zoom installer not found. Downloading first..."
        Download-Zoom
    }

    Write-Log "Installing Zoom silently..."
    Start-Process "msiexec.exe" -ArgumentList "/i `"$InstallerPath`" /quiet /norestart" -Wait

    if (Test-Path $ZoomPath) {
        Write-Log "Zoom installation completed successfully."
        Remove-Item -Path $InstallerPath -Force
    } else {
        Write-Log "ERROR: Zoom installation failed."
    }
}

# Function to uninstall Zoom silently
function Uninstall-Zoom {
    Write-Log "Searching for installed Zoom versions..."

    # Check for Zoom in the registry
    $ZoomRegPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $ZoomEntries = @()
    foreach ($RegPath in $ZoomRegPaths) {
        $ZoomEntries += Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue | 
            Where-Object { $_.DisplayName -match "Zoom Workplace" }
    }

    if ($ZoomEntries.Count -eq 0) {
        Write-Log "No installed versions of Zoom found."
        return
    }

    foreach ($Entry in $ZoomEntries) {
        Write-Log "Uninstalling Zoom version: $($Entry.DisplayName)"

	Write-Output $Entry.UninstallString

        if ($Entry.UninstallString) {
            # Extract MSI product code if applicable
            if ($Entry.UninstallString -match "/X\{(.+?)\}") {
                $ProductCode = "{$($matches[1])}"
                Start-Process "msiexec.exe" -ArgumentList "/x $ProductCode /quiet /norestart" -Wait
                Write-Log "Uninstalled Zoom using MSI Product Code: $ProductCode"
            }
        } else {
            Write-Log "ERROR: No uninstall string found for $($Entry.DisplayName)"
        }
    }
}

# Menu for user selection
while ($true) {
    Clear-Host
    Write-Output "Zoom Management Menu"
    Write-Output "1. Download Zoom Installer"
    Write-Output "2. Install Zoom (Silent Mode)"
    Write-Output "3. Uninstall Zoom (All Versions)"
    Write-Output "4. Disable Firewall"
    Write-Output "5. Enable Firewall"
    Write-Output "6. Start Zoom Stress Load"
    Write-Output "7. Stop Zoom Stress Load"
    Write-Output "8. Exit"
    $Choice = Read-Host "Enter your choice (1-6)"

    switch ($Choice) {
        "1" { Download-Zoom }
        "2" { Install-Zoom }
        "3" { Uninstall-Zoom }
        "4" { Disable-Firewall }
        "5" { Enable-Firewall }
        "6" { Start-ZoomStressLoad }
        "7" { Stop-ZoomStressLoad }
        "8" { Write-Log "User exited the script."; exit }
        default { Write-Log "ERROR: Invalid choice entered ($Choice)." }
    }

    Pause
}
