# killer-scripts

A collection of production-ready PowerShell scripts for Windows administration, deployment, repair, and hardening. Most scripts require an elevated PowerShell session to function correctly. Scripts that enforce this automatically are noted in their headers. Most are compatible with RMM platforms for unattended execution.

**Compatibility:** PowerShell 5.1, PowerShell 7, and Kaseya LiveConnect. No external dependencies unless noted in the script header. All scripts are tested on Windows 10 and Windows 11.

Repository: <https://github.com/SteveTheKiller/killer-scripts>

---

## AMORT.ps1

### Advanced Maintenance, Optimization & Restoration Tool

Advanced Windows tune-up script designed for MSP field deployment and remote management support. Good first response when a client triggers a low disk space alert, a machine feels sluggish, or you need to standardize a freshly imaged Dell before handing it off.

Runs a sequence of hardening and optimization tasks in a single pass. The script collects system data via WMI/CIM (with fallback for PowerShell 7 compatibility), disables Windows AI features and Recall functionality, hardens privacy settings across all user profiles, strips Dell bloatware including SupportAssist and Dell Update, removes OEM manufacturer software, purges %TEMP%, C:\Windows\Temp, browser caches, and Windows Update delivery caches with MB-recovered reporting at each stage, resets the Windows Update database via DISM, executes DISM RestoreHealth and System File Checker repairs, and forces TRIM optimization on SSDs.

Key function Get-SystemData handles WMI to CIM compatibility by detecting PowerShell edition and using CIM for PS7 or WMI for earlier versions.

RMM/Unattended support enabled. Safe to execute on systems with active users without intervention required. Exit code 0 on success.

---

## BERET.ps1

### BitLocker Encryption, Recovery & Escrow Tool

BitLocker lifecycle manager for provisioning, monitoring, and recovery key escrow.

Interactive wizard for managing BitLocker across domain and non-domain systems. Detects TPM status and Secure Boot configuration, offers FIPS or Standard mode selection for new BitLocker deployments, initializes TPM when needed, supports Active Directory and Entra ID escrow for recovery keys, generates and displays recovery keys in the console, and enforces policy refresh via gpupdate.

Interactive UI uses colored prompts for user feedback. Detects domain membership and domain controller status to route escrow appropriately. Returns exit code 0 for RMM compatibility.

RMM/Unattended support requires pre-configuration or will prompt for selections. Suitable for enterprise key management workflows.

---

## DEBLOAT.ps1

### Deployment Environment Bloat Liquidator & Optimized Automated Toolkit

Windows 11 standardization script removing OEM manufacturer bloat, AI features, and telemetry.

Targets HP, Dell, ASUS, and Acer OEM software removal. Disables Windows Recall and AI-related features. Applies telemetry caps and enforces privacy settings. Hardens Microsoft Edge configuration. Cleans privacy and telemetry artifacts across all user profiles, including the Default User template, ensuring that new profiles created after execution inherit the hardened baseline.

Key function Invoke-ComprehensiveUserCleanup iterates all user hives (local and remote registry paths) to remove traces of telemetry services, advertising, tracking, and AI features at the per-user level.

RMM/Unattended support enabled. Processes all profiles atomically in a single execution. Exit code 0 on success.

---

## DEFEND.ps1

### Definition Enforcement & Full Endpoint Network Defense

Kernel-level Windows security auditing covering hardware, firmware, software, and threat detection.

Performs comprehensive security assessment of TPM, Secure Boot, HVCI (Hypervisor-protected Code Integrity), and Windows Defender hardening. Configures the Windows Firewall based on domain status (domain-joined machines receive more restrictive inbound rules; workgroup machines use less restrictive rules). Triggers Defender threat definitions synchronization. Executes a full malware scan with live IOPS reporting so the administrator can monitor disk activity and scan progress. Correlates event logs to identify threats and suspicious activity. Generates exit code 0 for RMM integration.

RMM/Unattended support enabled. Long-running due to full scan; suitable for off-hours scheduling. Detailed threat reporting for enterprise SOC integration.

---

## DEPOT.ps1

### Deployment & Endpoint Provisioning Operations Tool

New-machine provisioning script deploying common enterprise applications and hardening.

Automated deployment of Microsoft 365 (M365), Teams, OneDrive, Google Chrome, Adobe Acrobat Reader, Zoom, and 7-Zip. ESC-to-skip allows users to abort any step interactively. Triggers Windows Update automation via UsoClient. Applies privacy and UI hardening across all user profiles. Self-deletes the script after successful completion to clean up.

Designed for mass deployment via RMM or imaging systems. Returns exit code 0 on completion. Safe for unattended execution.

---

## FACTS.ps1

### Foxit Audit and Control Task Script

Foxit PDF reader installation auditing and automatic update blocking.

Detects Foxit PDF Reader installation. Blocks automatic updates by hardening registry keys and suppressing the Foxit update service. Creates a scheduled maintenance task running hourly to self-heal and re-enforce the update block, ensuring settings persist across reboots and manual changes.

Useful for organizations standardizing on Foxit while preventing unwanted automatic upgrades. Returns exit code 0 for RMM compatibility.

RMM/Unattended support enabled.

---

## MACE.ps1

### Microsoft Application Cleanse & Eradication

Complete removal of OneDrive, New Outlook, Office, M365, Microsoft Project, and Teams.

Interactive wizard allowing per-user or system-wide removal scope selection. Fully uninstalls Microsoft Office suite, M365, Teams, and Project. Removes OneDrive integration and cleans shell folder registry entries to undo OneDrive redirect policies. Flushes cached credentials and authentication tokens to ensure clean state.

Function Get-TargetedUsers identifies which user profiles to process based on scope selection. Restores shell folder paths (Documents, Desktop, Downloads) to local locations if they were redirected to OneDrive.

RMM/Unattended support requires pre-configuration of scope. Interactive UI otherwise. Exit code 0 on completion.

---

## ODD.ps1

### Output Device Diagnostic

Audio device inventory with categorization and configuration reporting.

Enumerates all audio input and output devices on the system. Categorizes devices as physical (integrated audio), USB audio, or Bluetooth audio. Reports the currently active default device with highlighting. Queries device driver version, sample rate, bit depth, and exclusive mode configuration from WMI and registry.

Useful for audit and troubleshooting audio hardware and driver state. RMM-compatible output for asset tracking.

---

## ORCA.ps1

### Outlook Repair & Configuration Assistant

Outlook repair utility for both New Outlook (Microsoft Store) and Classic Outlook (Office/M365).

Interactive wizard with repair scope selection. Resets New Outlook (Microsoft Store edition) by clearing application state, cache, and extensions. Resets Classic Outlook (Office or M365 edition) by deleting OST files, clearing cache, removing mail extensions, scrubbing registry keys, and flushing authentication tokens. Reinstalls New Outlook from the Microsoft Store after reset if it was the active version.

Function Invoke-OutlookReset is the core repair engine. Includes OST file backup before deletion. Clears COM add-ins and mail client associations to remove corrupt extensions.

RMM/Unattended support requires pre-configuration. Interactive UI otherwise. Exit code 0 on completion.

---

## PRINT.ps1

### Printer Response & Interface Network Tool

Printer management, network discovery, and driver installation.

Multi-threaded network scanning discovers printers on the local network via port 9100 (LPD) and port 631 (CUPS). Scrapes printer HTTP/Web server interfaces to extract device information and recommended drivers. Intelligently matches local driver stores to discovered devices. Falls back to IPP Class Driver if manufacturer driver unavailable. Supports manual IP entry and UNC shared printer paths.

Interactive UI with table output of discovered printers, driver availability, and installation status. Core function Invoke-MultiThreadedNetworkScan uses PowerShell runspace threading for parallel port scanning.

RMM/Unattended support requires network configuration pre-setup. Useful for imaging and fleet printing standardization.

---

## PRUNE.ps1

### Profile Removal Utility for Neglected Entries

User profile staleness analyzer and orphaned profile detector.

Scans all local user profiles and reports staleness indicators including LastUseTime, total folder size, and orphaned or disabled account status. Flags profiles older than 90 days (stale) and older than 365 days (old). Allows interactive profile selection for deletion or direct username targeting via -Username parameter.

Investigates why live user profiles remain open by querying services and scheduled tasks that may be holding file handles. Function Get-ProfileStaleness walks the user profile registry hive and file timestamps.

RMM-compatible output for reports. Exit code 0 on completion.

---

## SHADE.ps1

### System Hardening Against Data Exposure

Comprehensive privacy hardening removing telemetry, tracking, and advertising.

Disables Windows telemetry services DiagTrack and dmwappushservice. Removes the Advertising ID. Blocks camera and microphone access for built-in apps. Disables Activity History. Disables clipboard logging (prevents clipboard data synchronization to cloud). Disables application telemetry feedback collection. Disables Delivery Optimization (P2P download optimization). Hardens network behavior to prevent phone-home activity by built-in applications.

Configured for Windows Pro, Business, and Enterprise SKUs. Does not modify Home edition (licensing restrictions). Returns exit code 0 for RMM compatibility.

RMM/Unattended support enabled. Safe for all supported SKUs. Suitable for privacy-sensitive deployments.

---

## STARE.ps1

### Scheduled Task Administration & Routine Executor

Interactive Scheduled Task creation wizard with trigger, command, and script execution support.

Guides users through creation of new scheduled tasks with daily, weekly, monthly, or startup triggers. Supports task execution of direct PowerShell commands or selection of scripts from the filesystem via an interactive browser with 'cd [path]' navigation for deep folder traversal.

Driver selection menu allows choosing which printer driver to use for devices discovered during PRINT.ps1 execution. INF file path can be specified manually for custom driver installation.

COM-based registration used for Monthly task triggers because New-ScheduledTaskTrigger cmdlet is unavailable in PowerShell 7. Interactive UI with color-coded prompts.

RMM/Unattended support requires pre-configuration of all task parameters. Otherwise interactive.

---

## TICK.ps1

### Trigger Immediate Clock Kickstart

Windows Time service synchronization resync and NTP peer validation.

Forces the Windows Time service to synchronize against an NTP peer. Detects domain membership and domain controller to offer domain-aligned NTP defaults. Interactive NTP server selection from 7 default NIST time servers or manual entry. Reports before/after timestamps with stratum level, root dispersion, leap indicator, and clock health summary.

Verifies time synchronization is working correctly and that the system clock is within acceptable bounds. Useful for troubleshooting time-dependent authentication and certificate validation failures.

RMM-compatible output for compliance reporting. Exit code 0 on success.

---

## URT.ps1

### Universal Rename Tool

Local or domain-joined computer rename utility preserving AD trust relationships.

Supports renaming computers on non-domain systems or domain-joined systems. For domain-joined machines, detects whether the machine is a workstation or domain controller and applies appropriate rename logic. Includes inline domain credential collection for systems requiring domain privilege to execute the rename. Optionally forces immediate reboot after rename.

Function Test-DomainRole returns the machine's role within AD. Preserves existing trust relationships during domain rename operations.

RMM/Unattended support requires credentials pre-staged or will prompt interactively. Exit code 0 on completion.

---

## VITALS.ps1

### Visual Interface for Technical Asset & Logistics Summary

Hardware and network inventory snapshot with visual formatting.

Collects and displays a full technical asset snapshot including manufacturer, model, serial number, CPU cores and clock speed, memory capacity and speed, GPU name and VRAM, disk health and BitLocker status, network interface configuration (IP, subnet, gateway, DNS), battery wear percentage and runtime estimate, TPM version and status, Secure Boot status, domain and Entra ID membership detection, and local administrator group membership.

VM environment detection identifies VMware, VirtualBox, Hyper-V, QEMU, Xen, and Parallels. Intelligent OEM filtering falls back to motherboard information when system manufacturer is generic or blank. BitLocker recovery key enumeration displays all recovery passwords for encrypted volumes.

Requires Administrator privilege. Colored table output formatted for 80-column terminal. Exit code 0 on completion. RMM-compatible for asset tracking and hardware audit reports.

---

## WURSA.ps1

### Windows Update, Repair & System Alignment

Windows Update, Repair, and System Alignment automation with feature upgrade support.

Enforces all essential and optional OS patches, OEM driver updates, and third-party application upgrades via Chocolatey. Intelligently skips applications currently in use to avoid disrupting the active user. Auto-installs Chocolatey if not present.

Includes support for Windows feature version upgrades with live heartbeat indicator showing installation progress, staged file size, and ESC-to-detach functionality. Safe for unattended and RMM execution via optional parameters.

Parameters -InplaceUpgrade auto-confirms feature upgrade prompt for unattended use. Parameter -No3rdParty skips the Chocolatey third-party app update pass entirely. Parameter -NoUpgrade skips featu
