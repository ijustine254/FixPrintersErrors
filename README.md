# Windows Printer Services Fix Script

A comprehensive PowerShell script designed to diagnose, repair, and optimize Windows printer services and configurations.

## Overview

`FixPrinterServices.ps1` is a PowerShell script that automates the process of troubleshooting and fixing common printer-related issues in Windows environments. The script performs multiple maintenance tasks and configurations to ensure proper printer functionality.

## Features

- **Administrator Privilege Check**
  - Ensures the script runs with proper administrative permissions

- **Windows Features Management**
  - Enables required printing features:
    - PDF Services
    - LPD (Line Printer Daemon) Service
    - LPR (Line Printer Remote) Port Monitor

- **Print Spooler Maintenance**
  - Clears stuck print jobs
  - Resets the print spooler service
  - Removes temporary printer files
  - Configures automatic startup

- **Registry Configurations**
  - Sets up RPC connection settings
  - Configures printer driver isolation
  - Creates missing registry paths if needed

- **Network Printer Diagnostics**
  - Tests connectivity to network printers
  - Verifies printer port configurations
  - Reports printer accessibility status

- **Print Job Management**
  - Removes stuck print jobs
  - Cleans up printer queues
  - Resets problematic print jobs

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or higher
- Administrator privileges
- Network connectivity (for network printer tests)

## Usage

1. Right-click on `FixPrinterServices.ps1` and select "Run with PowerShell as Administrator"
   
   OR

2. Open PowerShell as Administrator and run:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
   .\FixPrinterServices.ps1
   ```

## Output

The script provides detailed feedback including:
- Status of each operation
- Error messages if issues occur
- Connectivity test results
- Final printer configuration status

## Common Issues Fixed

- Stuck print jobs
- Print spooler service issues
- Network printer connectivity problems
- Registry configuration errors
- Print queue problems
- Driver isolation issues

## Safety Features

- Error handling for all critical operations
- Registry path verification before modifications
- Service status verification
- Automatic spooler service recovery
- Timeout limits for network operations

## Notes

- Always ensure you have a system backup before running system maintenance scripts
- Some operations require network connectivity
- The script may require a system restart to complete all changes
- Certain antivirus software may need to be temporarily disabled

## Contributing

Feel free to submit issues and enhancement requests via GitHub issues.

## License

[MIT License](LICENSE)

## Author

[Your Name/Organization]

## Last Updated

[Current Date] 