# Check for Administrator privileges
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script requires Administrator privileges. Please run as Administrator."
    exit 1
}

# Function to test and create registry paths
function Test-AndCreateRegistryPath {
    param([string]$Path)
    if (!(Test-Path $Path)) {
        try {
            New-Item -Path $Path -Force | Out-Null
            Write-Host "Created registry path: $Path"
            return $true
        } catch {
            Write-Error "Failed to create registry path $Path : $_"
            return $false
        }
    }
    return $true
}

# Enable Required Windows Features
Write-Host "Enabling Printer and Document Services..."
$features = @(
    'Printing-PrintToPDFServices-Features',
    'Printing-LPDPrintService',
    'Printing-LPRPortMonitor'
)

foreach ($feature in $features) {
    try {
        $state = Get-WindowsOptionalFeature -Online -FeatureName $feature
        if ($state.State -ne 'Enabled') {
            Enable-WindowsOptionalFeature -Online -FeatureName $feature -All -NoRestart -ErrorAction Stop
            Write-Host "Enabled feature: $feature"
        } else {
            Write-Host "Feature already enabled: $feature"
        }
    } catch {
        Write-Error "Failed to enable feature $feature : $_"
        exit 1
    }
}

# Configure Registry Settings for RPC Connections
Write-Host "Configuring RPC connection settings..."
$rpcPath = "HKLM:\Software\Policies\Microsoft\Windows NT\Printers\RPC"
if (Test-AndCreateRegistryPath $rpcPath) {
    try {
        Set-ItemProperty -Path $rpcPath -Name RpcUseNamedPipeProtocol -Type DWord -Value 1 -ErrorAction Stop
        Write-Host "RPC settings configured successfully"
    } catch {
        Write-Error "Failed to set RPC settings: $_"
        exit 1
    }
}

# Function to clear printer queue and reset spooler
function Reset-PrintQueue {
    Write-Host "Clearing print queue and temporary files..."
    try {
        $spoolPath = "$env:SystemRoot\System32\spool\PRINTERS"
        
        # Stop the spooler first
        Stop-Service -Name Spooler -Force -ErrorAction Stop
        
        # Wait for service to fully stop
        Start-Sleep -Seconds 2
        
        # Clear printer queue files if they exist
        if (Test-Path $spoolPath) {
            Get-ChildItem -Path $spoolPath -File | Remove-Item -Force -ErrorAction Stop
        }
        
        # Start spooler again
        Start-Service -Name Spooler -ErrorAction Stop
        Write-Host "Print queue cleared successfully"
    } catch {
        Write-Error "Failed to clear print queue: $_"
        # Ensure spooler is started even if cleanup fails
        Start-Service -Name Spooler -ErrorAction SilentlyContinue
        exit 1
    }
}

# Function to repair common printer issues
function Repair-PrinterIssues {
    Write-Host "Checking for common printer issues..."
    try {
        # Check and fix printer driver isolation
        $printPath = "HKLM:\System\CurrentControlSet\Control\Print"
        if (Test-AndCreateRegistryPath $printPath) {
            Set-ItemProperty -Path $printPath -Name "RpcAuthnLevelPrivacyEnabled" -Type DWord -Value 0 -ErrorAction Stop
        }
        
        # Reset Internet printing components
        Write-Host "Resetting Internet printing components..."
        $webClient = Get-Service -Name WebClient -ErrorAction SilentlyContinue
        if ($webClient) {
            if ($webClient.Status -eq 'Running') {
                Restart-Service -Name WebClient -Force -ErrorAction SilentlyContinue
            } else {
                Start-Service -Name WebClient -ErrorAction SilentlyContinue
            }
        }
        
        # Reset printer ports using CIM instead of WMI
        Write-Host "Resetting printer ports..."
        $ports = Get-CimInstance -ClassName Win32_TCPIPPrinterPort
        foreach ($port in $ports) {
            try {
                $port | Set-CimInstance -ErrorAction Stop
            } catch {
                Write-Warning "Failed to reset port $($port.Name): $_"
            }
        }
        
        # Check for stuck print jobs using CIM
        $printJobs = Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop
        if ($printJobs) {
            Write-Host "Removing stuck print jobs..."
            foreach ($job in $printJobs) {
                try {
                    $job | Remove-CimInstance -ErrorAction Stop
                } catch {
                    Write-Warning "Failed to remove print job: $_"
                }
            }
        }
        
    } catch {
        Write-Error "Failed to repair printer issues: $_"
        exit 1
    }
}

# Perform maintenance
Write-Host "Performing printer maintenance..."
Reset-PrintQueue
Repair-PrinterIssues

# Test printer connectivity with timeout
Write-Host "Testing printer connectivity..."
try {
    $printers = Get-Printer -ErrorAction Stop
    foreach ($printer in $printers) {
        if ($printer.PortName -match "^IP_") {
            $ipAddress = ($printer.PortName -split "_")[1]
            Write-Host "Testing connection to printer $($printer.Name) at $ipAddress..."
            
            # Test with timeout
            $result = Test-Connection -ComputerName $ipAddress -Count 1 -TimeoutSeconds 5 -Quiet
            if ($result) {
                Write-Host "Printer $($printer.Name) is accessible at $ipAddress" -ForegroundColor Green
            } else {
                Write-Warning "Printer $($printer.Name) is not responding at $ipAddress"
            }
        }
    }
} catch {
    Write-Warning "Failed to test printer connectivity: $_"
}

# Final Status Check
Write-Host "`nPerforming final status check..."
try {
    $spooler = Get-Service -Name Spooler
    Write-Host "Spooler Service Status: $($spooler.Status)" -ForegroundColor $(if ($spooler.Status -eq 'Running') { 'Green' } else { 'Red' })
    
    $printerCount = (Get-Printer).Count
    Write-Host "Total Printers Configured: $printerCount"
} catch {
    Write-Warning "Failed to get final status: $_"
}

# Recap and Final Steps
Write-Host "`nConfiguration complete. Please verify printer names and network settings to ensure proper operation."
