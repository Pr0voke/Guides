# Intune Autopilot Guide

## Gather Hash

First you'll need to make sure the computer is connected to the internet, either by going through enough setup steps to select a network, pluggin in an ethernet cable, or opening settings and using the GUI.

Once Connected you'll need to plug in a USB Drive containing the *GetHash.ps1* script.

1. Open a CMD window and type *powershell*

<img width="1920" height="1200" alt="1" src="https://github.com/user-attachments/assets/8f8a3761-536d-4815-9b95-02caad82bc27" />
![CMD Powershell](1.png)

2. Run the *GetHash.ps1* script by entering **D:\GetHash.ps1** in the powershell window
<details>
  <summary>Automatic Powershell Script, save this as GetHash.ps1</summary>
  
  ```
<#
.SYNOPSIS
    Collects the Windows Autopilot hardware hash (4K Hardware ID) from the local PC
    and uploads it directly to Intune via Microsoft Graph.

.DESCRIPTION
    Run this on a NEW machine from an elevated PowerShell session (or from the OOBE
    shell: press Shift+F10 at the setup screen).

    The script will:
      1. Read the 4K Hardware ID from WMI (MDM_DevDetail_Ext01).
      2. Export it to a CSV as a backup / for manual import.
      3. Authenticate to Microsoft Graph and import the device into Autopilot.
      4. Wait for the import + profile assignment to finish (optional).

.PARAMETER GroupTag
    Autopilot group tag (order ID). Hard-coded default is set in the param block
    below, so technicians do not need to type it. Only pass this if you need to
    override the standard tag for a one-off machine.

.PARAMETER AssignedUser
    UPN to pre-assign the device to. Leave blank for pre-provisioning / shared use.

.PARAMETER OutputFile
    Where to write the CSV backup. Defaults to a USB-friendly path next to the script.

.PARAMETER CsvOnly
    Skip the upload and just produce the CSV (for bulk import in the Intune portal).

.PARAMETER NoWait
    Skip polling. By default the script waits for the import to complete and for an
    Autopilot profile to be assigned before it exits.

.EXAMPLE
    .\Upload-AutopilotHash.ps1 -CsvOnly -OutputFile D:\Hashes\device.csv

.NOTES
    Required Graph permission: DeviceManagementServiceConfig.ReadWrite.All
    Signing in requires an account with the Intune Administrator role (or equivalent).
#>

[CmdletBinding()]
param(
    [string]$GroupTag = 'Haynie-Dell',
    [string]$AssignedUser,
    [string]$OutputFile = (Join-Path $PSScriptRoot "AutopilotHWID_$env:COMPUTERNAME.csv"),
    [switch]$CsvOnly,
    [switch]$NoWait
)

$ErrorActionPreference = 'Stop'

function Write-Step { param($Message) Write-Host "[*] $Message" -ForegroundColor Cyan }
function Write-Ok   { param($Message) Write-Host "[+] $Message" -ForegroundColor Green }
function Write-Warn { param($Message) Write-Host "[!] $Message" -ForegroundColor Yellow }

#region Pre-flight -----------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    throw "This script must be run from an elevated PowerShell session."
}

# TLS 1.2 for gallery / Graph calls on clean OOBE images
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
#endregion

#region Collect the hardware hash --------------------------------------------
Write-Step "Reading Hardware ID from WMI..."

$session = New-CimSession
$devDetail = Get-CimInstance -CimSession $session `
    -Namespace 'root/cimv2/mdm/dmmap' `
    -Class 'MDM_DevDetail_Ext01' `
    -Filter "InstanceID='Ext' AND ParentID='./DevDetail'"

if (-not $devDetail) {
    throw "Could not retrieve the hardware hash. This machine may not support Autopilot registration, or WMI is unavailable."
}

$serial = (Get-CimInstance -CimSession $session -Class Win32_BIOS).SerialNumber
$model  = (Get-CimInstance -CimSession $session -Class Win32_ComputerSystem).Model
$hash   = $devDetail.DeviceHardwareData
Remove-CimSession $session

Write-Ok "Serial: $serial  |  Model: $model"
#endregion

#region CSV backup ------------------------------------------------------------
$record = [ordered]@{
    'Device Serial Number' = $serial
    'Windows Product ID'   = ''
    'Hardware Hash'        = $hash
}
if ($GroupTag)     { $record['Group Tag']     = $GroupTag }
if ($AssignedUser) { $record['Assigned User'] = $AssignedUser }

$parent = Split-Path -Parent $OutputFile
if ($parent -and -not (Test-Path $parent)) { New-Item -Path $parent -ItemType Directory -Force | Out-Null }

[pscustomobject]$record | Export-Csv -Path $OutputFile -NoTypeInformation -Encoding UTF8
Write-Ok "CSV written to: $OutputFile"

if ($CsvOnly) {
    Write-Warn "CsvOnly specified - skipping upload. Import this file at: Intune > Devices > Enrollment > Devices > Import."
    return
}
#endregion

#region Connect to Graph ------------------------------------------------------
Write-Step "Checking for the Microsoft.Graph.Authentication module..."
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Write-Warn "Module not found. Installing from PSGallery (current user scope)..."
    Get-PackageProvider -Name NuGet -ForceBootstrap | Out-Null
    Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
}
Import-Module Microsoft.Graph.Authentication -Force

Write-Step "Signing in to Microsoft Graph (use an Intune Administrator account)..."
Connect-MgGraph -Scopes 'DeviceManagementServiceConfig.ReadWrite.All' -NoWelcome
#endregion

#region Import into Autopilot -------------------------------------------------
$uri  = 'https://graph.microsoft.com/v1.0/deviceManagement/importedWindowsAutopilotDeviceIdentities'
$body = @{
    serialNumber        = $serial
    hardwareIdentifier  = $hash
    groupTag            = $GroupTag
    assignedUserPrincipalName = $AssignedUser
} | ConvertTo-Json -Depth 3

Write-Step "Uploading hardware hash to Intune..."
try {
    $import = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $body -ContentType 'application/json'
}
catch {
    throw "Upload failed: $($_.Exception.Message)"
}

Write-Ok "Import submitted. Id: $($import.id)  |  State: $($import.state.deviceImportStatus)"
#endregion

#region Optional wait ---------------------------------------------------------
if (-not $NoWait) {
    Write-Step "Waiting for import to complete (this usually takes 1-5 minutes)..."
    $deadline = (Get-Date).AddMinutes(20)
    do {
        Start-Sleep -Seconds 20
        $status = Invoke-MgGraphRequest -Method GET -Uri "$uri/$($import.id)"
        $state  = $status.state.deviceImportStatus
        Write-Host "    import status: $state"
    } while ($state -eq 'unknown' -and (Get-Date) -lt $deadline)

    if ($state -eq 'complete') {
        Write-Ok "Device imported into Autopilot successfully."
    }
    elseif ($state -eq 'error') {
        Write-Warn "Import error $($status.state.deviceErrorCode): $($status.state.deviceErrorName)"
    }
    else {
        Write-Warn "Timed out waiting for import. Check Intune > Devices > Enrollment > Devices."
    }

    # Wait for a deployment profile to land on the device
    Write-Step "Waiting for an Autopilot profile to be assigned..."
    $devUri  = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$serial')"
    $deadline = (Get-Date).AddMinutes(20)
    do {
        Start-Sleep -Seconds 30
        $dev = (Invoke-MgGraphRequest -Method GET -Uri $devUri).value | Select-Object -First 1
        Write-Host "    profile status: $($dev.deploymentProfileAssignmentStatus)"
    } while ($dev.deploymentProfileAssignmentStatus -notlike 'assigned*' -and (Get-Date) -lt $deadline)

    if ($dev.deploymentProfileAssignmentStatus -like 'assigned*') {
        Write-Ok "Profile assigned. The device is ready to reboot into Autopilot."
    }
    else {
        Write-Warn "Profile not assigned yet. Verify the device is in the right dynamic group / group tag."
    }
}
#endregion

Disconnect-MgGraph | Out-Null
Write-Ok "Done. Reboot the PC, then press the Windows key 5 times at the sign-in screen to start pre-provisioning."

  ```

</details>

![Script](2.png)

3. Wait for this to finish and it will upload the hash automatically and assign the group. 

![Script Finish](3.png)

4. Once you've verified the Hash is upload in intune restart the computer 
>It will show up in Intune Admin>Devices>Enrollment>Windows AutoPilot>Devices

![Devices](4.png)

5. In the CMD powershell window type **shutdown /r /t 0**

![Shutdown](5.png)

## Run Autopilot Provisioning

1. Once the computer reboots press the **Windows key 5 times** to enter the OOBE Autopilot provisioning screen.

![Reboot OOBE](6.png)

2. On this screen select **Pre-provision with Windows Autopilot** and hit next.

![Pre-Provisioning](7.png)

3. This will run for about 30 min and install the required applications and 

![Autopilot Running](8.png)

4. Once it's finished select **Reseal** and it will shut the computer down and you're good to assign a user or give this to the user to sign in as normal

![Reseal](9.png)


## Reset Device after Termination or reassignment

1. Find the device in Intune.

![Find](10.png)

2. Select **Remove data>Autopilot Reset** Check the box and select **Action**
> This process will take anywhere from an hour to 2 hours to complete and requires the device to be connected to the internet. The reset is triggered by a restart if not left alone for it to automatically start after 45 min.

![Reset](11.png)

![Action](12.png)

## Alternate Hash script to CSV
> Requires manual upload and group assignment to the Enrollment>Devices
<details>

<summary>Get Hash to CSV Script</summary>

```
# ============================================================
# Get Intune Autopilot Hardware Hash
# Save/Append to CSV on USB with correct Intune headers
# ============================================================

$CsvName = "AutopilotHWID.csv"

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host "      INTUNE AUTOPILOT HASH COLLECTION" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------
# Find removable USB drive
# ------------------------------------------------------------

Write-Host "Looking for USB drive..." -ForegroundColor Cyan

$UsbDrive = Get-CimInstance Win32_LogicalDisk |
    Where-Object {
        $_.DriveType -eq 2 -and
        $_.DeviceID
    } |
    Select-Object -First 1

if (-not $UsbDrive) {
    Write-Host "ERROR: No removable USB drive detected." -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

$CsvPath = Join-Path $UsbDrive.DeviceID $CsvName

Write-Host "USB detected: $($UsbDrive.DeviceID)" -ForegroundColor Green
Write-Host "CSV location: $CsvPath" -ForegroundColor Cyan
Write-Host ""

# ------------------------------------------------------------
# Configure PowerShell
# ------------------------------------------------------------

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

Set-ExecutionPolicy `
    -Scope Process `
    -ExecutionPolicy Bypass `
    -Force

$ScriptPath = "C:\Program Files\WindowsPowerShell\Scripts"

if ($env:Path -notlike "*$ScriptPath*") {
    $env:Path += ";$ScriptPath"
}

# ------------------------------------------------------------
# Install Get-WindowsAutopilotInfo if needed
# ------------------------------------------------------------

$AutopilotScript = Get-Command `
    Get-WindowsAutopilotInfo.ps1 `
    -ErrorAction SilentlyContinue

if (-not $AutopilotScript) {

    Write-Host "Installing Get-WindowsAutopilotInfo..." -ForegroundColor Yellow

    try {

        Install-Script `
            -Name Get-WindowsAutopilotInfo `
            -Force `
            -Confirm:$false `
            -ErrorAction Stop

        $AutopilotScript = Get-Command `
            Get-WindowsAutopilotInfo.ps1 `
            -ErrorAction Stop

        Write-Host "Get-WindowsAutopilotInfo installed." -ForegroundColor Green
    }
    catch {

        Write-Host "" 
        Write-Host "ERROR installing Get-WindowsAutopilotInfo:" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red

        Read-Host "Press Enter to exit"
        exit 1
    }
}

# ------------------------------------------------------------
# Create temporary CSV
# ------------------------------------------------------------

$TempCsv = Join-Path $env:TEMP "AutopilotHashTemp.csv"

if (Test-Path $TempCsv) {
    Remove-Item $TempCsv -Force
}

Write-Host "Collecting Autopilot hardware hash..." -ForegroundColor Cyan

try {

    # Generate a clean temporary Autopilot CSV
    & $AutopilotScript.Source `
        -OutputFile $TempCsv

    if (-not (Test-Path $TempCsv)) {
        throw "Autopilot script did not create the temporary CSV."
    }

    # --------------------------------------------------------
    # Import generated data
    # --------------------------------------------------------

    $AutopilotData = Import-Csv -Path $TempCsv

    if (-not $AutopilotData) {
        throw "No Autopilot information was returned."
    }

    # --------------------------------------------------------
    # Force EXACT Intune Autopilot header names
    # --------------------------------------------------------

    $CleanData = foreach ($Device in $AutopilotData) {

        [PSCustomObject][ordered]@{
            "Device Serial Number" = $Device.'Device Serial Number'
            "Windows Product ID"   = $Device.'Windows Product ID'
            "Hardware Hash"        = $Device.'Hardware Hash'
        }
    }

    # --------------------------------------------------------
    # Validate data before saving
    # --------------------------------------------------------

    foreach ($Device in $CleanData) {

        if ([string]::IsNullOrWhiteSpace($Device.'Device Serial Number')) {
            throw "Device Serial Number is blank."
        }

        if ([string]::IsNullOrWhiteSpace($Device.'Hardware Hash')) {
            throw "Hardware Hash is blank."
        }
    }

    # --------------------------------------------------------
    # Create or append USB CSV
    # --------------------------------------------------------

    if (Test-Path $CsvPath) {

        Write-Host "Existing Autopilot CSV found." -ForegroundColor Yellow

        # Import existing CSV so we can validate it
        $ExistingData = Import-Csv -Path $CsvPath -ErrorAction Stop

        # Rebuild all data using correct headers
        $CombinedData = @()

        foreach ($Device in $ExistingData) {

            $CombinedData += [PSCustomObject][ordered]@{
                "Device Serial Number" = $Device.'Device Serial Number'
                "Windows Product ID"   = $Device.'Windows Product ID'
                "Hardware Hash"        = $Device.'Hardware Hash'
            }
        }

        # Prevent duplicate serial numbers
        foreach ($NewDevice in $CleanData) {

            $Duplicate = $CombinedData |
                Where-Object {
                    $_.'Device Serial Number' -eq
                    $NewDevice.'Device Serial Number'
                }

            if ($Duplicate) {

                Write-Host ""
                Write-Host "Device already exists in CSV:" -ForegroundColor Yellow
                Write-Host $NewDevice.'Device Serial Number' -ForegroundColor Yellow
            }
            else {

                $CombinedData += $NewDevice

                Write-Host ""
                Write-Host "Adding device:" -ForegroundColor Green
                Write-Host $NewDevice.'Device Serial Number' -ForegroundColor Green
            }
        }

        # Rewrite CSV completely with clean headers
        $CombinedData |
            Export-Csv `
                -Path $CsvPath `
                -NoTypeInformation `
                -Encoding UTF8 `
                -Force
    }
    else {

        Write-Host "Creating new Autopilot CSV..." -ForegroundColor Cyan

        $CleanData |
            Export-Csv `
                -Path $CsvPath `
                -NoTypeInformation `
                -Encoding UTF8 `
                -Force
    }

    # --------------------------------------------------------
    # Verify resulting CSV
    # --------------------------------------------------------

    $Verify = Import-Csv -Path $CsvPath

    $ExpectedHeaders = @(
        "Device Serial Number"
        "Windows Product ID"
        "Hardware Hash"
    )

    $ActualHeaders = @(
        $Verify[0].PSObject.Properties.Name
    )

    $HeadersValid =
        ($ActualHeaders.Count -eq $ExpectedHeaders.Count) -and
        (@(Compare-Object $ExpectedHeaders $ActualHeaders).Count -eq 0)

    Write-Host ""

    if (-not $HeadersValid) {

        Write-Host "ERROR: CSV headers failed validation." -ForegroundColor Red

        Write-Host "Headers detected:" -ForegroundColor Yellow

        $ActualHeaders | ForEach-Object {
            Write-Host "  $_"
        }

        throw "Generated CSV does not contain the correct Intune Autopilot headers."
    }

    # --------------------------------------------------------
    # Success
    # --------------------------------------------------------

    Write-Host "==============================================" -ForegroundColor Green
    Write-Host " AUTOPILOT HASH COLLECTED SUCCESSFULLY" -ForegroundColor Green
    Write-Host "==============================================" -ForegroundColor Green

    Write-Host ""
    Write-Host "Serial Number:" -ForegroundColor Cyan
    Write-Host $CleanData[0].'Device Serial Number' -ForegroundColor White

    Write-Host ""
    Write-Host "Saved to:" -ForegroundColor Cyan
    Write-Host $CsvPath -ForegroundColor White

    Write-Host ""
    Write-Host "Total devices in CSV: $($Verify.Count)" -ForegroundColor Cyan

    Write-Host ""
    Write-Host "CSV Headers:" -ForegroundColor Cyan

    Write-Host "Device Serial Number | Windows Product ID | Hardware Hash" `
        -ForegroundColor Green
}
catch {

    Write-Host ""
    Write-Host "ERROR collecting Autopilot hardware hash:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}
finally {

    if (Test-Path $TempCsv) {
        Remove-Item $TempCsv -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ""
Read-Host "Press Enter to exit"
```

</details>





