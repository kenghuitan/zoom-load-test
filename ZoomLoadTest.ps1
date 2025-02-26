<#
    Author  : Benjamin TAN
    Date    : 26 Feb 2025
    Purpose : Automate Zoom installation, stress test, and cleanup with logging and CPU/memory failsafe.
#>

# Get current script directory
$CurrentPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$DownloadURL = "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64"
$InstallerPath = "$CurrentPath\ZoomInstaller.msi"
$ZoomPath = "C:\Program Files\Zoom\bin\zoom.exe"  # Update if necessary
$ZoomDir = "C:\Program Files\Zoom\bin"  # Folder containing zoom executables
$LogFile = "$CurrentPath\ZoomAutomation.log"  # Log file in the same directory as script

# Function to write logs with timestamps
function Write-Log {
    param ([string]$Message)
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogEntry = "$Timestamp - $Message"
    Add-Content -Path $LogFile -Value $LogEntry
    Write-Output $LogEntry
}

# Function to start Zoom stress load with CPU and memory failsafe
function Start-ZoomStressLoad {
    if (-Not (Test-Path $ZoomPath)) {
        Write-Log "ERROR: Zoom.exe not found at $ZoomPath. Please install Zoom first."
        return
    }

    $MeetingID = Read-Host "Enter the Meeting ID"
    $MeetingCode = Read-Host "Enter the Meeting Code"
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
    
    # Find and stop all running Zoom stress instances
    $ZoomProcesses = Get-Process | Where-Object { $_.ProcessName -match "^zoom\d+$" }
    if ($ZoomProcesses) {
        $ZoomProcesses | ForEach-Object { Stop-Process -Id $_.Id -Force }
        Write-Log "All Zoom stress test instances have been stopped."
    } else {
        Write-Log "No Zoom stress test instances found."
    }

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

# Function to download Zoom installer to the same folder as the script
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
    $Choice = Read-Host "Enter your choice (1-5)"

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
