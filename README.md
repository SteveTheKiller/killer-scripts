<div align="center">

[![Killer Scripts](header.png)](https://killertools.net/killer-scripts)

</div>

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%20%7C%207-5391FE?logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6?logo=windows&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-blue)
![Scripts](https://img.shields.io/badge/Scripts-19-brightgreen)

A collection of PowerShell scripts for Windows administration, deployment, repair, and hardening. Each script is a single self-contained file with no module dependencies.

**Compatibility:** Windows PowerShell 5.1, PowerShell 7, and Kaseya LiveConnect and other RMM consoles, including runs as SYSTEM. Tested on Windows 10 and Windows 11. Every script except BEAM and ORCA requires an elevated session, and DEPOT relaunches itself elevated.

**License:** MIT. See [LICENSE](LICENSE).

---

## Exit codes

Every script uses the same exit codes, and its final status line always agrees with the code it returns.

| Code | Meaning |
| --- | --- |
| 0 | Success |
| 1 | Failure |
| 2 | Not elevated, a precondition failed, or the run was cancelled. Nothing was changed. |
| 3 | Partial success. Some steps worked and some did not. |
| 3010 | Success, reboot required |

Interactive scripts never hang without a console. They either refuse with exit 2 or take a safe default, as noted in each section.

---

## Contents

| Script | Description |
| --- | --- |
| [AMORT](#amortps1) | Disk cleanup and Windows repair for MSP field and remote support |
| [BEAM](#beamps1) | Cloud-only Intune Win32 app deployer for pushing any installer to Entra-joined devices |
| [BERET](#beretps1) | BitLocker lifecycle manager with TPM enforcement and AD / Entra ID key escrow |
| [DEBLOAT](#debloatps1) | Windows 11 standardization removing OEM bloat, AI features, and telemetry |
| [DEFEND](#defendps1) | Security audit of TPM, Secure Boot, HVCI, Defender, and the firewall |
| [DEPOT](#depotps1) | New-machine provisioning with silent app deployment and hardening |
| [FACTS](#factsps1) | Foxit install audit and automatic update blocking |
| [MACE](#maceps1) | Clean-slate removal of OneDrive, New Outlook, Office, M365, Project, and Teams |
| [ODD](#oddps1) | Read-only audio device inventory with default device and driver reporting |
| [ORCA](#orcaps1) | Outlook reset for New (Store) and Classic (Office / M365) editions |
| [PEEK](#peekps1) | Live read-only monitor for Windows Update and Setup servicing activity |
| [PRINT](#printps1) | Printer management, network discovery, and driver installation |
| [PRUNE](#pruneps1) | User profile staleness analyzer and safe profile removal |
| [SHADE](#shadeps1) | Privacy hardening against telemetry, tracking, and advertising |
| [STARE](#stareps1) | Scheduled Task browser and creation wizard |
| [TICK](#tickps1) | Windows Time service resync with source and last-sync verification |
| [URT](#urtps1) | Local or domain-joined computer rename |
| [VITALS](#vitalsps1) | Read-only hardware, security, and network inventory snapshot |
| [WURSA](#wursaps1) | Windows Update, driver, feature upgrade, and third-party app enforcement |

---

## AMORT.ps1

### Advanced Maintenance, Optimization & Restoration Tool

Disk cleanup and Windows repair for MSP field deployment and remote support. A good first response to a low disk space alert or a sluggish machine.

Runs eleven steps in one pass:

- Removes Delivery Optimization cache, finished upgrade leftovers, OEM driver extract folders, and Dell SupportAssist remediation snapshots on Dell hardware.
- Removes crash dumps, Windows Error Reporting files, and Recycle Bin items older than the retention window, and temp files older than two days.
- Clears browser, Teams, and GPU shader caches.
- Clears the Windows Update download cache. The update database itself is kept.
- Trims System Restore shadow copies to the newest few. Other shadow copies are never touched.
- Compresses log folders, runs DISM RestoreHealth, component cleanup, and SFC, flushes DNS, turns hibernation off, and retrims SSDs.

User documents, mail, sync caches, and search indexes are never touched. DISM and SFC are skipped while a reboot is pending or servicing is active. Windows.old and the upgrade rollback window are kept unless you ask for their removal.

| Parameter | Effect |
| --- | --- |
| `-DryRun` | Reports what would be removed without changing anything |
| `-RetainDays <1-365>` | Age threshold for dumps, error reports, and the Recycle Bin. Default 14. |
| `-KeepRestorePoints <1-64>` | Restore points to keep. Default 2. |
| `-RemoveRollback` | Also removes Windows.old and the upgrade rollback files. Pasted runs can set `$RemoveRollback = $true` instead. |

Never prompts, so it is safe for unattended RMM runs. Handles a 32-bit host through Sysnative and the 64-bit registry view. Exit codes 0, 1, 2, 3, and 3010.

---

## BEAM.ps1

### Bulk Endpoint App Mover

Cloud-only Intune Win32 app deployment tool for pushing any installer to Entra-joined devices.

Packages a local or hosted .exe or .msi as an Intune Win32 app and deploys it to one or more Entra-joined devices as SYSTEM, with no LAN access and no stored credentials. It prompts for the installer source (a URL, or Enter to browse for a file), app name, publisher, optional install switches, target hostnames, and an admin sign-in. MSI packages default to `/qn`. Sign-in uses a browser prompt with PKCE and a loopback listener, so nothing gets pasted.

Before packaging, BEAM shows the source URL, SHA256, and signer. A link scraped from a download page needs an explicit yes. The Microsoft Win32 Content Prep Tool is pinned to v1.8.7 and its signature is verified before use.

Each run creates its own Entra group, and each target must match exactly one device record before it is added. The install writes a per-app marker containing the run ID only when the installer exits 0, 3010, or 1641, so a failed install is never reported as detected. Offline devices install on their next Intune check-in.

Runs from the technician's workstation, not the endpoint, and is interactive only. Without a console it exits 2. Exit codes 0, 1, 2, and 3. The first use in a new tenant needs one-time admin consent to the Graph permissions.

---

## BERET.ps1

### BitLocker Encryption, Recovery & Escrow Tool

BitLocker lifecycle manager for provisioning, repair, and recovery key escrow.

- Requires Pro or Enterprise. Home exits 2.
- FIPS mode uses XtsAes256 and sets the FIPS algorithm policy. Standard mode uses XtsAes128.
- Readies the TPM when needed, including Dell BIOS activation. If the TPM needs a reboot first, it exits 3010.
- A volume encrypted with the wrong cipher is decrypted and re-encrypted, and data volume auto-unlock is suspended and restored around the cycle. A volume with the right cipher but no recovery key gets one added in place.
- Displays the recovery keys on screen for the technician.
- Escrows the keys to Active Directory, Entra ID or a Microsoft account, and an optional key file. Each destination is read back and tagged VERIFIED, UNCHECKED, or FAILED.

| Preset | Effect |
| --- | --- |
| `$BeretMode = 'FIPS'` or `'STANDARD'` | Skips the mode prompt |
| `$BeretKeyFile = '<path>'` | Also writes the recovery passwords to this file and reads it back |

Without a console and without `$BeretMode`, it exits 2 and changes nothing. If no escrow destination succeeds, it exits 3 so the key never silently goes unrecorded. Handles a 32-bit host through Sysnative. Exit codes 0, 1, 2, 3, and 3010.

---

## DEBLOAT.ps1

### Deployment Environment Bloat Liquidator & Optimized Automated Toolkit

Windows 11 standardization script removing OEM bloat, AI features, and telemetry.

- Removes Recall and sets the Windows AI data analysis policy.
- Disables DiagTrack and dmwappushservice, caps telemetry, and turns off error reporting, CEIP, the advertising ID, activity feed, consumer features, Spotlight, Chat, Feeds, and Copilot.
- Applies Edge policies for the sidebar, Copilot, background mode, startup boost, and shopping. Edge bookmarks and saved data are left alone.
- Sets the High performance power plan with sleep and display timeouts, and turns hibernation off.
- Removes HP, Dell, ASUS, Acer, and Lenovo OEM software, plus common third-party apps such as McAfee, Norton, Avast, AVG, TikTok, Netflix, and Spotify. The McAfee folder is kept when a McAfee or Trellix product is actually installed. Rivet Networks drivers are kept.
- Purges non-English Office language packs.
- Applies per-user settings to the Default User template, loaded hives, and offline hives, so new profiles inherit the baseline. Every hive it loads is unloaded again.

Never prompts, so it is safe for unattended RMM runs. Exit codes 0, 2, 3, and 3010. It returns 3010 when Recall removal needs a reboot.

---

## DEFEND.ps1

### Definition Enforcement & Full Endpoint Network Defense

Windows security audit and Defender hardening covering hardware, firmware, software, and threats.

- Enables real-time protection, PUA protection, cloud protection, network protection, and Controlled Folder Access, then reads real-time protection back to confirm it held.
- Audits VT-x, Virtual Machine Platform, Secure Boot, TPM, HVCI, and tamper protection.
- Lists Defender exclusions. Only the Huntress install folders are treated as known-good, and anything else is reported as UNVERIFIED. Exclusions are reported only, never changed.
- Enables the Domain, Private, and Public firewall profiles and reads them back. Network categories are left exactly as Windows set them.
- Updates definitions and checks the result.
- Runs a full scan only in an interactive session, where Esc skips it and Q switches to a quick scan. A scan counts as complete only when Defender's exit code and completion event both confirm it. Unattended runs skip the scan.
- Reports only active threats. Cleaned or quarantined history does not count.

Exit codes 0, 2, 3, and 3010. Active threats, refused settings, failed updates, or an unverified scan give exit 3. A pending reboot on an otherwise clean run gives 3010. Handles a 32-bit host through the 64-bit registry view.

---

## DEPOT.ps1

### Deployment & Endpoint Provisioning Operations Tool

New-machine provisioning script deploying common applications and hardening.

- Optional steps, each a y/N prompt: rename the computer, create a local admin, and remove other admins while disabling the built-in Administrator. The admin purge runs only after the new admin was created and verified in the same run, and it never removes the account running the script.
- Triggers Windows Update through UsoClient.
- Installs Microsoft 365, OneDrive, Teams, Chrome, Acrobat Reader, Zoom, and 7-Zip, skipping anything already installed. Downloads are size and header checked. With a console, Esc skips the current app.
- Force-installs uBlock Origin in Edge and uBlock Origin Lite in Chrome.
- Applies privacy and UI hardening across every user profile.
- Writes a run log to `C:\ProgramData\DEPOT\LastRun.txt` and deletes itself only after a run that ended 0 or 3010.

Without a console, every prompt answers N, so the rename and admin steps never run unattended. Relaunches itself elevated through UAC when needed. A normal run exits 3010, because the Windows Update step always flags a reboot. Exit codes 0, 1, 2, 3, and 3010.

---

## FACTS.ps1

### Foxit Audit and Control Task Script

Foxit installation audit and automatic update blocking.

When any Foxit product is installed, it sets the update-block registry values on every Foxit updater key in both registry views, disables the Foxit update services and scheduled tasks, and reads each change back. It then registers an hourly SYSTEM task, `FACTS_Universal_Enforcement`, that re-applies the same block, so the settings survive Foxit's own updater. Updating the task never deletes the working one first. When no Foxit product is found, it removes that task. It never uninstalls Foxit.

Never prompts, so it is safe for unattended RMM runs. Handles a 32-bit host through the 64-bit registry view. Exit codes 0, 1, 2, and 3.

---

## MACE.ps1

### Microsoft Application Cleanse & Eradication

Clean-slate removal of OneDrive, New Outlook, Office, M365, Microsoft Project, and Teams, including their settings.

- Uninstalls OneDrive, New Outlook, Office and M365 (Click-to-Run with an MSI fallback), Project, and Teams, including provisioned copies when targeting all users.
- Wipes the Microsoft app data and registry settings those apps leave behind, so a reinstall starts fresh.
- Flushes the cached Office and Teams credentials for the account running the script.
- Repairs a redirected shell folder only when its target is missing and no copy exists in another OneDrive folder.

MACE never deletes user data. Documents, Desktop, Pictures, the OneDrive sync folders, and every .pst file are preserved, and a kept .pst is reported so the run ends as exit 3 rather than claiming a full wipe.

| Preset | Effect |
| --- | --- |
| `$MaceScope = 'AllUsers'` | Targets every profile. This is the default. |
| `$MaceScope = '<profile folder>'` | Targets one profile under `C:\Users`. An unknown name exits 2. |

Interactively, pick a profile number, or press A or Enter for All Users. Esc aborts with nothing changed. Without a console and without a preset, it runs for All Users. Exit codes 0, 1, 2, and 3.

---

## ODD.ps1

### Output Device Diagnostic

Read-only audio device inventory.

Lists every input and output endpoint in two groups, Physical and USB, and Bluetooth, with status and driver version and date. The default device comes from Windows Core Audio and is matched by device identity, then highlighted. Sample rate, bit depth, and exclusive mode are shown for the default endpoint. If Core Audio cannot answer, it falls back to a registry guess labelled "(inferred)". It also reports the Windows Audio service state.

Changes nothing, but needs elevation. Never prompts. Handles a 32-bit host through the 64-bit registry view. Exit codes 0, 2, and 3. A stopped audio service or a device in an error state gives exit 3, and a disabled device does not.

---

## ORCA.ps1

### Outlook Repair & Configuration Assistant

Outlook reset for both New Outlook (Microsoft Store) and Classic Outlook (Office / M365), for one profile or all of them.

- **New Outlook:** removes the app, clears its data, and reinstalls it from the Microsoft Store. For all users it downloads the package before removing anything.
- **Classic Outlook:** clears the roaming cache, offline address books, logging, and the AutoDiscover and general option keys.
- **OST files:** optional. They are renamed to `.ost.bak-<timestamp>`, never deleted, and the rename is verified.
- **Add-ins:** optional. Only the Outlook add-in registry keys are removed.
- **Sign-out:** clears the Office identity and OneAuth caches and the Outlook and Office entries in Credential Manager. Credential Manager can only be cleared for the account running the script, so other profiles are reported as a problem.

Only the selected user's Outlook processes are stopped. Without elevation it works on the current profile only.

| Parameter | Effect |
| --- | --- |
| `-Silent` | No prompts. All users when elevated, otherwise the current user. Resets both editions without renaming OSTs or purging add-ins. |
| `-FullSignOut` | Also clears the Edge WebView, Token Broker, and AAD Broker caches. This signs the user out of more than Outlook. |

Without a console and without `-Silent`, it exits 2. Exit codes 0, 1, 2, and 3.

---

## PEEK.ps1

### Patch & Event Examination Kit

Read-only live monitor for Windows Update and Windows Setup servicing activity, built for remote support triage.

Correlates Windows Update client events, active servicing processes, Panther setup logs, compatibility blocks, reboot flags, service state, disk space, power and BitLocker state, and OS build into one refreshing console view. A setup.exe process is only counted as servicing when other evidence backs it up, and a compatibility file it cannot parse is not reported as a block.

| Parameter | Effect |
| --- | --- |
| `-RefreshSeconds <1-60>` | Refresh interval. Default 3. |
| `-Once` | Renders one snapshot and exits |

Press Q or Esc to quit. Use `-Once` for unattended runs, because the live loop has nothing to stop it without a console. Makes no changes, starts no updates, and never reboots. Requires elevation. Exit codes 0, 1, and 2.

---

## PRINT.ps1

### Printer Response & Interface Network Tool

Interactive printer management, network discovery, and driver installation.

- Adds a printer by network scan, manual IP, or UNC share path, and deletes installed printers after a confirmation.
- The scan probes TCP 9100 (RAW / JetDirect) and 631 (IPP) on the subnet of the default route. A network larger than /22 is limited to the /24 around this machine.
- Reads the model from each printer's web page and matches it against the local driver store, falling back to the Microsoft IPP Class Driver.
- Installs over RAW 9100 or IPP, and installs INF drivers through pnputil.

| Parameter | Effect |
| --- | --- |
| `-Silent` | Lists the installed printers and exits. Never adds or removes anything. |

Without a console and without `-Silent`, it exits 2. Requires elevation. Exit codes 0, 1, 2, and 3.

---

## PRUNE.ps1

### Profile Removal Utility for Neglected Entries

User profile staleness analyzer and safe profile removal.

Lists every local profile, oldest first, with last use, folder size, and account status. Profiles unused for over 90 days are flagged stale, and over 365 days old. Status is matched by SID: Local, Disabled, Domain, Orphaned (a local SID that no longer resolves), Unresolved (a domain or Entra SID that could not be looked up, never treated as orphaned), or LookupFailed.

Removal takes a Y/N and then typing DELETE. Each profile is re-checked right before deletion, and loaded, special, or signed-in profiles are refused. A leftover folder is only removed when it sits directly under the profiles root, is not a link, and no other profile owns it. The result is verified afterward.

When every profile is loaded, it lists the services and scheduled tasks running as each account, which is usually why the profile stays open.

| Parameter | Effect |
| --- | --- |
| `-Username <name>` | Targets one profile. The confirmations still apply. |
| `-SkipSizes` | Skips the folder size scan |

Without a console it cannot confirm, so it exits 2 and removes nothing. Exit codes 0, 1, 2, and 3.

---

## SHADE.ps1

### System Hardening Against Data Exposure

Privacy hardening against telemetry, tracking, and advertising, in eight steps.

1. **Location:** turns off location and sensors and disables the geolocation service.
2. **Telemetry:** sets telemetry to the minimum and turns off error reporting, CEIP, and inventory collection. Disables DiagTrack, WerSvc, and PcaSvc. dmwappushservice is deliberately left running, because MDM push and device sync depend on it.
3. **Advertising:** turns off the advertising ID, tailored experiences, suggested content, Bing and Cortana in search, and device metadata downloads.
4. **Camera and microphone:** set to Deny in the app consent store. The policy is left in the user's control.
5. **History:** turns off Activity History, clipboard history and sync, program tracking, and input personalization.
6. **Feedback and Spotlight:** turns off feedback prompts, online speech, Find My Device, and Spotlight content.
7. **Delivery Optimization:** sets peer downloads off, sets the service to Manual, and purges its cache.
8. **Network:** turns off active probing, LLMNR, and the peer networking features, and sets NetBIOS to point-to-point.

Per-user values are written to every signed-in user's hive when it runs as SYSTEM. Users who sign in later need another run. With nobody signed in, the per-user values are skipped and the run ends as exit 3.

Never prompts, so it is safe for unattended RMM runs. A reboot is recommended afterward. Exit codes 0, 2, and 3.

---

## STARE.ps1

### Scheduled Task Administration & Routine Executor

Interactive Scheduled Task browser and creation wizard.

Lists every task outside `\Microsoft\`, showing its folder, state, and next run. Deleting asks for a confirmation, targets the task by folder and name so a same-named task elsewhere is safe, and verifies the result.

The wizard creates a task that runs either a program with arguments or a .ps1 picked from a file browser, where `cd [path]` jumps to a folder. Triggers are daily, weekly, monthly (any days, or 0 for the last day), or at startup. Every task runs as SYSTEM with highest privileges. Replacing an existing task updates it in place, so a failed replacement leaves the working task alone. If the chosen script can be edited by non-administrators, STARE warns and asks before scheduling it, because that would let anyone with write access run code as SYSTEM.

Without a console it exits 2 and changes nothing. Requires elevation. Exit codes 0, 1, 2, and 3.

---

## TICK.ps1

### Trigger Immediate Clock Kickstart

Windows Time service reset and resync with verification.

Re-registers and restarts the Windows Time service, configures the peer, and forces a resync, checking every step. Domain members can sync from the domain hierarchy. Otherwise the interactive menu offers Cloudflare, NIST, pool.ntp.org, Microsoft, Google, Apple, Ubuntu, or a custom server. It reports the source, last sync, stratum, root dispersion, and leap indicator.

The run passes only when the reported source matches the one configured and the last successful sync is under ten minutes old.

| Parameter | Effect |
| --- | --- |
| `-Silent` | No menu. Uses the domain hierarchy on a domain member, otherwise time.cloudflare.com. |
| `-NtpServer <host>` | Uses this server, without the menu. The name is validated first. |

Without a console it runs as if `-Silent` was given. Handles a 32-bit host through Sysnative. Exit codes 0, 1, 2, and 3. A sync that was configured but cannot be confirmed gives exit 3.

---

## URT.ps1

### Universal Rename Tool

Interactive computer rename for workgroup, Entra-joined, and domain-member machines.

Prompts for the new name. On a domain member it also prompts for domain credentials and renames with them, so the computer account is renamed in AD. Domain controllers are refused with exit 2 and pointed to `netdom computername`. A blank name, Esc, or the current name exits 0 with nothing changed.

After a rename, typing RESTART reboots immediately and exits 0. Anything else leaves the reboot to you and exits 3010.

Without a console it exits 2 and changes nothing. Requires elevation. Exit codes 0, 1, 2, and 3010.

---

## VITALS.ps1

### Visual Interface for Technical Asset & Logistics Summary

Read-only hardware, security, and network inventory snapshot.

Shows device name, uptime, last boot, public IP, domain and Entra join state, VM detection, TPM, Secure Boot, manufacturer, model, and serial (falling back to the motherboard when the OEM fields are blank), OS build, CPU, memory, GPU, battery wear and runtime, physical disks with health and BitLocker state, active network adapters, local administrators, and a pending reboot flag.

The BitLocker section lists every key protector on every volume, including volumes without a drive letter, with its type and ID. Recovery password protectors also show the key.

The public IP lookup is an outbound request to ifconfig.me with a two-second timeout. Never prompts. Requires elevation. Exit codes 0, 2, and 3. Exit 3 means the report ran but the disks or BitLocker could not be read.

---

## WURSA.ps1

### Windows Update, Repair & System Alignment

Windows Update, driver, feature upgrade, and third-party app enforcement in five steps.

1. Creates a restore point with a time limit. A failure here does not stop the run.
2. Registers the Microsoft Update service so Office and other product updates come through.
3. Installs every pending software and driver update through the Windows Update API, one at a time, resuming a paused Windows Update. A pending reboot defers this pass.
4. When a feature upgrade may be needed, temporarily opens the Windows Update policy and rolls it back afterward, unless the upgrade now owns it.
5. Updates third-party apps through Chocolatey, but only apps already installed from a fixed list. Apps in use are skipped. Chocolatey is installed only after you agree, and its installer must be signed by Chocolatey Software before it runs.

The feature upgrade targets 26H2. It runs as a detached SYSTEM task that survives the session ending, and it reports to `C:\Windows\Temp\25H2IPU\ipu_status.txt`. It picks the enablement package, Windows Update, the Installation Assistant, or a hash-verified ISO depending on the build, edition, and language. LTSC is skipped. It refuses on battery power. WURSA never reboots the machine itself.

| Parameter | Effect |
| --- | --- |
| `-InplaceUpgrade` | Approves the feature upgrade and the third-party pass without prompting |
| `-Update3rdParty` | Approves the third-party pass without prompting |
| `-No3rdParty` | Skips the third-party pass |
| `-NoUpgrade` | Skips the feature upgrade check |

Without a console every prompt answers no. On a machine below 26H2, pass either `-InplaceUpgrade` or `-NoUpgrade` for unattended runs. Otherwise the upgrade is declined and the run exits 1 as "FEATURE UPGRADE NOT COMPLETE". Exit codes 0, 1, 2, 3, and 3010.
