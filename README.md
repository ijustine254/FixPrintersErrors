# FixPrinterServices.ps1

This PowerShell script is designed to diagnose and fix common printer-related issues on Windows systems. It automates tasks such as restarting the Print Spooler service, clearing the print queue, resetting printer ports, and repairing registry settings. The script also includes robust error handling and logging for easy troubleshooting.

---

## Features

- **Restart Print Spooler Service**: Stops and restarts the Print Spooler service to resolve stuck print jobs.
- **Clear Print Queue**: Clears all pending print jobs and temporary files in the spooler folder.
- **Enable Required Windows Features**: Enables essential printer-related Windows features (e.g., LPD, LPR, Print to PDF).
- **Repair Printer Issues**: Resets printer ports, fixes driver isolation settings, and removes stuck print jobs.
- **Rebuild Spooler Folder**: Clears and rebuilds the spooler folder to resolve corruption issues.
- **Reinstall Printer Drivers**: Reinstalls printer drivers to fix driver-related issues.
- **Reset Network Settings**: Resets the TCP/IP stack and flushes DNS to resolve network printer connectivity issues.
- **Reset Printer Permissions**: Resets printer permissions to default settings.
- **Logging**: Logs all actions and errors to a file for easy debugging.
- **Execution Policy Handling**: Temporarily sets the execution policy to `Bypass` to ensure the script runs, then restores the original policy afterward.

---

## Prerequisites

- **PowerShell**: The script requires PowerShell 5.1 or later.
- **Administrator Privileges**: The script must be run as an administrator.
- **Windows Features**: Ensure that the required Windows features (e.g., Print and Document Services) are available on your system.

---

## Usage

1. **Download the Script**:
   - Save the script as `FixPrinterServices.ps1`.

2. **Run the Script**:
   - Open PowerShell as an administrator.
   - Navigate to the directory where the script is saved.
   - Run the script:
     ```powershell
     .\FixPrinterServices.ps1
     ```

3. **Review the Logs**:
   - The script logs all actions and errors to a file located at:
     ```
     %TEMP%\PrinterFixLog.txt
     ```
   - Check this file for detailed output and troubleshooting information.

---

## Script Workflow

1. **Temporarily Set Execution Policy**:
   - The script sets the execution policy to `Bypass` to ensure it runs without restrictions.

2. **Check for Administrator Privileges**:
   - The script verifies that it is running with administrator privileges.

3. **Perform Printer Maintenance**:
   - The script performs the following tasks in sequence:
     - Restart the Print Spooler service.
     - Enable required Windows features.
     - Configure RPC connection settings.
     - Clear the print queue.
     - Repair common printer issues.
     - Rebuild the spooler folder.
     - Reinstall printer drivers.
     - Reset network settings.
     - Reset printer permissions.

4. **Restore Execution Policy**:
   - The script restores the original execution policy after completing all tasks.

---