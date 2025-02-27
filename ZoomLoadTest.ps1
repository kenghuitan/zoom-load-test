<#
    Author  : Benjamin TAN
    Date    : 26 Feb 2025
    Purpose : Automate Zoom installation, stress test, and cleanup with logging and CPU/memory failsafe.
#>

# Get current script directory
$CurrentPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$DownloadURL = "https://zoom.us/client/latest/ZoomInstaller.exe?archType=x64"
$InstallerPath = "$CurrentPath\ZoomInstaller.exe"
$ZoomPath = "C:\Program Files\Zoom\bin\zoom.exe"  # Update if necessary
$ZoomDir = "C:\Program Files\Zoom\bin"  # Folder containing zoom executables
$LogFile = "$CurrentPath\ZoomAutomation.log"
$ConfigFile = "$CurrentPath\ZoomConfig.json"

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
        if ($UseConfig -match "^[Yy]$") {
            $MeetingID = $Config.MeetingID
            $MeetingCode = $Config.MeetingCode
        }
    }

    # If no saved values or user wants to enter new ones
    if (-not $MeetingID -or -not $MeetingCode) {
        $MeetingID = Read-Host "Enter the Meeting ID"
        $MeetingCode = Read-Host "Enter the Meeting Code"
        Save-Config -MeetingID $MeetingID -MeetingCode $MeetingCode
    }

    $RunCount = Read-Host "Enter the number of times to copy and run Zoom.exe"
    if ($RunCount -match '^\d+$') {
        $RunCount = [int]$RunCount  
        Write-Log "Starting Zoom stress test with $RunCount instances (Meeting ID: $MeetingID, Code: $MeetingCode)."

        for ($i = 1; $i -le $RunCount; $i++) {
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
            $NewZoomPath = "$ZoomDir\zoom$i.exe"
            Copy-Item -Path $ZoomPath -Destination $NewZoomPath -Force
            Write-Log "Copied zoom.exe to zoom$i.exe"

            # Start the copied Zoom instance with the meeting link
            $MeetingURL = "zoommtg://zoom.us/join?action=join&uname=User$i&confno=$MeetingID&pwd=$MeetingCode&opt=join&role=0"
            Start-Process -FilePath $NewZoomPath -ArgumentList "--url=`"$MeetingURL`"" -NoNewWindow
            Write-Log "Started zoom$i.exe with meeting link: $MeetingURL"
        }
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
}

# Function to download Zoom installer
function Download-Zoom {
    Write-Log "Starting Zoom download to $InstallerPath..."

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

# Menu for user selection
while ($true) {
    Clear-Host
    Write-Output "Zoom Management Menu"
    Write-Output "1. Download Zoom Installer"
    Write-Output "2. Install Zoom (Silent Mode)"
    Write-Output "3. Start Zoom Stress Load"
    Write-Output "4. Stop Zoom Stress Load"
    Write-Output "5. Exit"
    $Choice = Read-Host "Enter your choice (1-6)"

    switch ($Choice) {
        "1" { Download-Zoom }
        "2" { Install-Zoom }
        "3" { Start-ZoomStressLoad }
        "4" { Stop-ZoomStressLoad }
        "5" { Write-Log "User exited the script."; exit }
        default { Write-Log "ERROR: Invalid choice entered ($Choice)." }
    }

    Pause
}
