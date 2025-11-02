# Save the current execution policy
$originalExecutionPolicy = Get-ExecutionPolicy -Scope Process

# Temporarily set the execution policy to Bypass to allow script execution
Set-ExecutionPolicy Bypass -Scope Process -Force
Write-Host "Execution policy temporarily set to Bypass for this script." -ForegroundColor Yellow

# Check for Administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Host "This script requires Administrator privileges. Please run as Administrator." -ForegroundColor Red
    # Restore the original execution policy before exiting
    Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
    exit 1
}

# Logging setup
$logFile = "$env:TEMP\PrinterFixLog.txt"
function Write-Log {
    param([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "[$timestamp] $message"
    Write-Host $message
}

# Function to test and create registry paths
function Test-AndCreateRegistryPath {
    param([string]$Path)
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
            Write-Log "Created registry path: $Path"
        }
        return $true
    } catch {
        Write-Log "ERROR: Failed to create registry path '$Path': $_"
        Write-Host "An error occurred while creating a registry path. Please check the log file for details." -ForegroundColor Red
        return $false
    }
}

# Function to restart the Spooler service
function Restart-SpoolerService {
    try {
        # Check if the Spooler service exists
        $spoolerService = Get-Service -Name Spooler -ErrorAction Stop
        Write-Log "Spooler service found. Restarting..."

        # Stop and start the service
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        Start-Sleep -Seconds 2
        Start-Service -Name Spooler -ErrorAction Stop
        Write-Log "Spooler service restarted successfully."
    } catch {
        if ($_ -like "*Cannot find any service with service name 'Spooler'*") {
            Write-Log "ERROR: The Spooler service was not found on this system."
            Write-Host "The Spooler service is not installed or is missing. Please ensure the Print Spooler feature is enabled." -ForegroundColor Red
        } else {
            Write-Log "ERROR: Failed to restart Spooler service: $_"
            Write-Host "An unexpected error occurred while restarting the Spooler service. Please check the log file for details." -ForegroundColor Red
        }
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to enable required Windows features
function Enable-PrinterFeatures {
    $features = @(
        'Printing-PrintToPDFServices-Features',
        'Printing-LPDPrintService',
        'Printing-LPRPortMonitor'
    )

    foreach ($feature in $features) {
        try {
            $state = Get-WindowsOptionalFeature -Online -FeatureName $feature -ErrorAction Stop
            if ($state.State -ne 'Enabled') {
                Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart -ErrorAction Stop
                Write-Log "Enabled feature: $feature"
            } else {
                Write-Log "Feature already enabled: $feature"
            }
        } catch {
            if ($_ -like "*The requested feature name is not valid*") {
                Write-Log "ERROR: The feature '$feature' was not found on this system."
                Write-Host "The feature '$feature' is not available on this version of Windows." -ForegroundColor Yellow
            } else {
                Write-Log "ERROR: Failed to enable feature '$feature': $_"
                Write-Host "An error occurred while enabling the feature '$feature'. Please check the log file for details." -ForegroundColor Red
            }
        }
    }
}

# Function to configure RPC connection settings
function Configure-RPCSettings {
    $rpcPath = "HKLM:\Software\Policies\Microsoft\Windows NT\Printers\RPC"
    try {
        if (Test-AndCreateRegistryPath $rpcPath) {
            Set-ItemProperty -Path $rpcPath -Name RpcUseNamedPipeProtocol -Type DWord -Value 0 -ErrorAction Stop
            Write-Log "RPC settings configured successfully."
        }
    } catch {
        Write-Log "ERROR: Failed to configure RPC settings: $_"
        Write-Host "An error occurred while configuring RPC settings. Please check the log file for details." -ForegroundColor Red
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to clear the print queue and reset the spooler
function Reset-PrintQueue {
    try {
        $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
        
        # Stop the spooler first
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        
        # Wait for service to fully stop
        Start-Sleep -Seconds 2
        
        # Clear printer queue files if they exist
        if (Test-Path $spoolPath) {
            Get-ChildItem -Path $spoolPath -File | Remove-Item -Force -ErrorAction Stop
            Write-Log "Cleared print queue files."
        }
        
        # Start spooler again
        Start-Service -Name Spooler -ErrorAction Stop
        Write-Log "Print queue cleared successfully."
    } catch {
        Write-Log "ERROR: Failed to clear print queue: $_"
        Write-Host "An error occurred while clearing the print queue. Please check the log file for details." -ForegroundColor Red
        # Ensure spooler is started even if cleanup fails
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to repair common printer issues
function Repair-PrinterIssues {
    try {
        # Check and fix printer driver isolation
        $printPath = "HKLM:\System\CurrentControlSet\Control\Print"
        if (Test-AndCreateRegistryPath $printPath) {
            Set-ItemProperty -Path $printPath -Name "RpcAuthnLevelPrivacyEnabled" -Type DWord -Value 0 -ErrorAction Stop
            Write-Log "Fixed printer driver isolation settings."
        }
        
        # Reset Internet printing components
        Write-Log "Resetting Internet printing components..."
        $webClient = Get-Service -Name WebClient -ErrorAction SilentlyContinue
        if ($webClient) {
            if ($webClient.Status -eq 'Running') {
                Restart-Service -Name WebClient -Force -ErrorAction SilentlyContinue
            } else {
                Start-Service -Name WebClient -ErrorAction SilentlyContinue
            }
            Write-Log "WebClient service reset."
        }
        
        # Reset printer ports
        Write-Log "Resetting printer ports..."
        $ports = Get-CimInstance -ClassName Win32_TCPIPPrinterPort
        foreach ($port in $ports) {
            try {
                $port | Set-CimInstance -ErrorAction Stop
                Write-Log "Reset printer port: $($port.Name)"
            } catch {
                Write-Log "WARNING: Failed to reset port $($port.Name): $_"
            }
        }
        
        # Check for stuck print jobs
        $printJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop
        if ($printJobs) {
            Write-Log "Removing stuck print jobs..."
            foreach ($job in $printJobs) {
                try {
                    $job | Remove-CimInstance -ErrorAction Stop
                    Write-Log "Removed print job: $($job.JobId)"
                } catch {
                    Write-Log "WARNING: Failed to remove print job: $_"
                }
            }
        }
    } catch {
        Write-Log "ERROR: Failed to repair printer issues: $_"
        Write-Host "An error occurred while repairing printer issues. Please check the log file for details." -ForegroundColor Red
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to rebuild the Spooler folder
function Rebuild-SpoolerFolder {
    try {
        $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        if (Test-Path $spoolPath) {
            Remove-Item -Path "$spoolPath\*" -Recurse -Force -ErrorAction Stop
            Write-Log "Cleared spooler folder."
        }
        Start-Service -Name Spooler -ErrorAction Stop
        Write-Log "Spooler folder rebuilt successfully."
    } catch {
        Write-Log "ERROR: Failed to rebuild spooler folder: $_"
        Write-Host "An error occurred while rebuilding the spooler folder. Please check the log file for details." -ForegroundColor Red
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to reinstall printer drivers
function Reinstall-PrinterDrivers {
    try {
        $drivers = Get-PrinterDriver -ErrorAction Stop
        foreach ($driver in $drivers) {
            Remove-PrinterDriver -Name $driver.Name -ErrorAction Stop
            Add-PrinterDriver -Name $driver.Name -ErrorAction Stop
            Write-Log "Reinstalled printer driver: $($driver.Name)"
        }
    } catch {
        Write-Log "ERROR: Failed to reinstall printer drivers: $_"
        Write-Host "An error occurred while reinstalling printer drivers. Please check the log file for details." -ForegroundColor Red
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to reset TCP/IP stack and flush DNS
function Reset-NetworkSettings {
    try {
        Write-Log "Resetting TCP/IP stack and flushing DNS..."
        netsh int ip reset | Out-Null
        ipconfig /flushdns | Out-Null
        Write-Log "Network settings reset successfully."
    } catch {
        Write-Log "ERROR: Failed to reset network settings: $_"
        Write-Host "An error occurred while resetting network settings. Please check the log file for details." -ForegroundColor Red
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Function to reset printer permissions
function Reset-PrinterPermissions {
    try {
        Write-Log "Resetting printer permissions..."
        $printers = Get-Printer -ErrorAction Stop
        foreach ($printer in $printers) {
            $printerName = $printer.Name
            $acl = Get-Acl -Path "HKLM:\System\CurrentControlSet\Control\Print\Printers\$printerName"
            $rule = New-Object System.Security.AccessControl.RegistryAccessRule("Everyone", "FullControl", "Allow")
            $acl.SetAccessRule($rule)
            Set-Acl -Path "HKLM:\System\CurrentControlSet\Control\Print\Printers\$printerName" -AclObject $acl
            Write-Log "Reset permissions for printer: $printerName"
        }
    } catch {
        Write-Log "ERROR: Failed to reset printer permissions: $_"
        Write-Host "An error occurred while resetting printer permissions. Please check the log file for details." -ForegroundColor Red
        # Restore the original execution policy before exiting
        Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
        exit 1
    }
}

# Main script logic
Write-Log "Starting printer maintenance..."
Restart-SpoolerService
Enable-PrinterFeatures
Configure-RPCSettings
Reset-PrintQueue
Repair-PrinterIssues
Rebuild-SpoolerFolder
Reinstall-PrinterDrivers
Reset-NetworkSettings
Reset-PrinterPermissions
Write-Log "Printer maintenance completed."
Write-Host "Printer maintenance completed. Please check the log file for details: $logFile" -ForegroundColor Green

# Restore the original execution policy
Set-ExecutionPolicy $originalExecutionPolicy -Scope Process -Force
Write-Host "Execution policy restored to $originalExecutionPolicy." -ForegroundColor Yellow
