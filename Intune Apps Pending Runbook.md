# Runbook: Intune Apps Stuck on "Pending" – Lost MDM Enrollment / Missing IME

**Applies to:** Entra-joined Windows devices where Intune apps sit on *Pending* indefinitely.
**Run all scripts in an elevated PowerShell window** (Run as administrator) unless noted.

---

## Symptoms

- Apps show **Pending** in Company Portal / Intune and never install.
- `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AppWorkload.log` repeats:
  `Failed to get device id, which will affect app status reports and device targeted policies`
- `IntuneManagementExtension.log` repeats:
  `Found 0 MDM certificates from Local Computer Store` / `Cert not found after retries`

**Root cause:** The device lost its Intune MDM certificate (enrollment broken), but a stale enrollment record remains, so Windows refuses to re-enroll ("device is already enrolled", `0x8018000A`). After re-enrolling, the Intune Management Extension (IME) may also need a clean reinstall (`0x80070643` / MSI 1603).

---

## Before you start

- [ ] Confirm the **BitLocker recovery key** is escrowed in Entra / Intune.
- [ ] Confirm the user has an **Intune license** and is in the **MDM user scope** (Entra → Mobility (MDM and WIP) → Microsoft Intune).
- [ ] Have **local admin** access on the device.
- [ ] The **user must be signed in** and the device **online** during enrollment.
- [ ] Create a temp folder for backups:

```powershell
New-Item -ItemType Directory -Path C:\Temp -Force | Out-Null
```

---

## Step 1 – Confirm the diagnosis

```powershell
# Join state – expect AzureAdJoined : YES and AzureAdPrt : YES
dsregcmd /status | Select-String "AzureAdJoined|AzureAdPrt|MdmUrl|DeviceId"

# Intune MDM certificate – nothing returned = cert missing
Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Issuer -like "*Intune MDM*" } |
  Select-Object Subject, NotBefore, NotAfter, Thumbprint

# IME service present?
Get-Service *Intune* | Select-Object Name, Status

# Recent "device id" failures in AppWorkload.log
Select-String -Path "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AppWorkload*.log" `
  -Pattern "Failed to get device id" | Select-Object -Last 3
```

If the cert is missing and AppWorkload.log shows `Failed to get device id`, continue.

---

## Step 2 – Identify the stale enrollment

```powershell
Get-ChildItem HKLM:\SOFTWARE\Microsoft\Enrollments | ForEach-Object {
  $p = Get-ItemProperty $_.PSPath
  [pscustomobject]@{GUID=$_.PSChildName; Provider=$p.ProviderID; UPN=$p.UPN; State=$p.EnrollmentState}
} | Where-Object Provider | Format-Table -AutoSize
```

**The stale enrollment** is the entry tied to the user's UPN, with Provider **`MS DM Server`** or **`Microsoft Device Management`**.

**Do NOT delete** these built-in entries: `WMI_Bridge_SCCM_Server`, `Deploy Authority`, `Cloud Authority`, `Local Authority`.

Optional sanity check (should point to `manage.microsoft.com`):

```powershell
$guid = "PASTE-STALE-GUID-HERE"
Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Enrollments\$guid" |
  Select-Object ProviderID, UPN, EnrollmentType, EnrollmentState, DiscoveryServiceFullURL
```

---

## Step 3 – Remove the stale enrollment

Also delete the stale device record in the Intune admin center (Devices → All devices) if one exists.

```powershell
$guid = "PASTE-STALE-GUID-HERE"

# Backup first
reg export "HKLM\SOFTWARE\Microsoft\Enrollments\$guid" "C:\Temp\Enrollment_$guid.reg" /y

$paths = @(
  "HKLM:\SOFTWARE\Microsoft\Enrollments\$guid",
  "HKLM:\SOFTWARE\Microsoft\Enrollments\Status\$guid",
  "HKLM:\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$guid",
  "HKLM:\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\$guid",
  "HKLM:\SOFTWARE\Microsoft\PolicyManager\Providers\$guid",
  "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$guid",
  "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Logger\$guid",
  "HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\$guid"
)
foreach ($p in $paths) { if (Test-Path $p) { Remove-Item $p -Recurse -Force; "Removed $p" } }

# Scheduled tasks + task folder for that enrollment
Get-ScheduledTask -TaskPath "\Microsoft\Windows\EnterpriseMgmt\$guid\" -ErrorAction SilentlyContinue |
  Unregister-ScheduledTask -Confirm:$false
$svc = New-Object -ComObject Schedule.Service; $svc.Connect()
try { $svc.GetFolder("\Microsoft\Windows\EnterpriseMgmt").DeleteFolder($guid,0); "Removed task folder" } catch {}
```

---

## Step 4 – Re-enroll via Group Policy

> Running `deviceenroller.exe /c /AutoEnrollMDM` manually did **not** work reliably (no event 75/76 logged). The local GPO method below is what worked.

1. Open `gpedit.msc` (as admin).
2. Go to **Computer Configuration → Administrative Templates → Windows Components → MDM**.
3. Open **Enable automatic MDM enrollment using default Azure AD credentials** → **Enabled**, Credential Type = **User Credential** → OK.
4. **Reboot**, sign in as the user (online).
5. Run:

```powershell
gpupdate /force
```

6. If enrollment doesn't start: **Task Scheduler → Microsoft → Windows → EnterpriseMgmt** → right-click **"Schedule created by enrollment client for automatically enrolling in MDM from AAD"** → **Run**.

### If gpupdate shows "Windows failed to apply the MDM Policy settings"

Get the real error code from event 1085:

```powershell
Get-WinEvent -LogName System -MaxEvents 300 |
  Where-Object Id -eq 1085 | Select-Object -First 1 | ForEach-Object {
    $x = [xml]$_.ToXml()
    $code = ($x.Event.EventData.Data | Where-Object Name -eq 'ErrorCode').'#text'
    $desc = ($x.Event.EventData.Data | Where-Object Name -eq 'ErrorDescription').'#text'
    "{0}  0x{1:X}  {2}" -f $_.TimeCreated, [int64]$code, $desc
  }
```

| Code | Meaning | Fix |
|---|---|---|
| `0x8018000A` | Device is already enrolled | Stale enrollment still present – repeat Steps 2–3 |
| `0x8018002B` | User not in MDM scope / UPN issue | Fix MDM user scope in Entra |
| `0x80180018` | No Intune license | Assign license |
| `0x80180001` | Server error / leftover enrollment | Repeat Steps 2–3 |
| `0xCAA…` / `0x82AA…` | Entra token issue | Reboot, sign in online as user |

---

## Step 5 – Verify enrollment and clean up extra certificates

```powershell
# Certificates, newest first
Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Issuer -like "*Intune MDM*" } |
  Select-Object Subject, NotBefore, NotAfter, Thumbprint | Sort-Object NotBefore -Descending

# Current Intune Device ID
Get-ChildItem HKLM:\SOFTWARE\Microsoft\Enrollments | ForEach-Object {
  $p = Get-ItemProperty "$($_.PSPath)\DMClient\MS DM Server" -ErrorAction SilentlyContinue
  if ($p.EntDMID) { "Enrollment $($_.PSChildName)  ->  Intune Device ID $($p.EntDMID)" }
}

# Enrollment events: 75 = success, 76 = failure
Get-WinEvent -LogName "Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin" |
  Where-Object Id -in 75,76 | Select-Object -First 3 | Format-List TimeCreated, Id, Message
```

- **Keep** the cert whose **CN matches the Intune Device ID**.
- If multiple enrollment attempts left orphaned certs, remove them:

```powershell
Remove-Item "Cert:\LocalMachine\My\<ORPHAN-THUMBPRINT>"
```

- In the Intune admin center, search the device name and **delete duplicate device records**, keeping only the one matching the current Intune Device ID.

---

## Step 6 – Restart or reinstall the IME

```powershell
Get-Service IntuneManagementExtension -ErrorAction SilentlyContinue
```

**If the service exists:**

```powershell
Restart-Service IntuneManagementExtension
```

Then skip to Step 7.

**If the service is missing**, sync (Settings → Accounts → Access work or school → Info → **Sync**), wait 10–15 min, and check the install status:

```powershell
Get-ChildItem "HKLM:\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement" -Recurse -ErrorAction SilentlyContinue |
  Get-ItemProperty | Where-Object { $_.CurrentDownloadUrl -like "*IntuneWindowsAgent*" } |
  Select-Object PSChildName, Status, @{n='LastErrorHex';e={'0x{0:X}' -f [int64]$_.LastError}}
```

- **Status 70** = installed.
- **Status 60 + `0x80070643`** (MSI 1603) = failed install → do 6a–6c.

### 6a – (Optional) Read the MSI failure reason

```powershell
Get-WinEvent -LogName Application -MaxEvents 300 |
  Where-Object { $_.ProviderName -eq "MsiInstaller" -and $_.Message -like "*Intune*" } |
  Select-Object -First 5 | Format-List TimeCreated, Id, Message
```

### 6b – Remove any leftover IME install

```powershell
$ime = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" |
  Where-Object DisplayName -like "*Intune Management Extension*"
$ime | Select-Object DisplayName, DisplayVersion, PSChildName
foreach ($i in $ime) { Start-Process msiexec.exe -ArgumentList "/x $($i.PSChildName) /qn /norestart" -Wait }

Stop-Process -Name AgentExecutor, Microsoft.Management.Services.IntuneWindowsAgent -Force -ErrorAction SilentlyContinue
Remove-Item "C:\Program Files (x86)\Microsoft Intune Management Extension" -Recurse -Force -ErrorAction SilentlyContinue
```

(Leave `C:\ProgramData\Microsoft\IntuneManagementExtension\Logs` in place.)

### 6c – Clear failed install records, reboot, sync

```powershell
Get-ChildItem "HKLM:\SOFTWARE\Microsoft\EnterpriseDesktopAppManagement" -Recurse -ErrorAction SilentlyContinue |
  Where-Object { (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).CurrentDownloadUrl -like "*IntuneWindowsAgent*" } |
  ForEach-Object { Remove-Item $_.PSPath -Recurse -Force; "Removed $($_.Name)" }
```

Reboot → sign in as the user → **Sync** → wait 10–15 min → rerun the status check above until **Status 70** and the service is running.

---

## Step 7 – Confirm apps are processing

```powershell
Restart-Service IntuneManagementExtension

# Should show no new "Failed to get device id" after the restart
Select-String -Path "C:\ProgramData\Microsoft\IntuneManagementExtension\Logs\AppWorkload.log" `
  -Pattern "Failed to get device id|Get policies|detection|download" | Select-Object -Last 10
```

In the Intune admin center, confirm the device shows a recent check-in. Pending apps should begin installing within ~1 hour.

---

## Step 8 – Optional cleanup

- Set the local GPO from Step 4 back to **Not Configured** if you don't want it left on the device.
- Delete `C:\Temp\Enrollment_*.reg` once the device is confirmed healthy.

---

## Quick reference – codes seen in this fix

| Value | Where | Meaning |
|---|---|---|
| `Failed to get device id` | AppWorkload.log | MDM cert missing → enrollment broken |
| Event **1085** | System log | MDM Policy (auto-enroll) failed – read ErrorCode |
| `2149056522` / `0x8018000A` | Event 1085 | Device already enrolled (stale record) |
| Event **75 / 76** | DeviceManagement-Enterprise-Diagnostics-Provider/Admin | Auto-enroll success / failure |
| Status **60**, `2147944003` / `0x80070643` | EnterpriseDesktopAppManagement | IME MSI install failed (1603) |
| Status **70** | EnterpriseDesktopAppManagement | IME installed |
