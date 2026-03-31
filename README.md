# killer-scripts

A collection of PowerShell tools built for MSP field technicians and systems administrators. Every script is designed to run reliably across **PowerShell 5.1, PowerShell 7, and Kaseya LiveConnect** with no external dependencies unless noted. All scripts require elevation.

Each script is named as an acronym — because if you're going to spend 40 hours writing a tool, it deserves a good name.

---

## Scripts

| Script | Full Name | Description |
|--------|-----------|-------------|
| **AMORT** | Advanced Maintenance, Optimization & Restoration Tool | Full Windows tune-up: AI/privacy hardening, OEM debloat, browser/system cache purge, Windows Update database reset, DISM RestoreHealth, SFC, and SSD TRIM. Reports disk space recovered at each stage. |
| **BERET** | BitLocker Encryption, Recovery & Escrow Tool | Interactive BitLocker lifecycle manager. Prompts for FIPS (AES-256) or Standard (AES-128) mode, initializes TPM, enforces compliance-gated encryption, generates recovery keys, and escrows to AD, Entra ID, or MSA depending on join state. |
| **DEBLOAT** | Deployment Environment Bloat Liquidator & Optimized Automated Toolkit | Standardizes Windows 11 by removing OEM bloat (HP, Dell, ASUS/Acer), AI/Recall features, and sponsored consumer content. Applies privacy hardening and taskbar/Start menu lockdown across all user profiles including the Default User template. |
| **DEFEND** | Definition Enforcement & Full Endpoint Network Defense | Audits and enforces kernel-level Windows security: TPM, Secure Boot, HVCI, and Defender hardening. Syncs threat definitions, runs a full Defender scan with live timer, and reports active threats with file paths pulled from event logs. |
| **DEPOT** | Deployment & Endpoint Provisioning Operations Tool | Full new-machine provisioning: silently deploys M365, Teams, OneDrive, Chrome, Acrobat Reader, Zoom, and 7-Zip with ESC-to-skip support. Triggers Windows Update, applies privacy/UI hardening across all profiles, and self-deletes on completion. |
| **FACTS** | Foxit Audit and Control Task Script | Audits for Foxit PDF installations and permanently blocks auto-updates via registry hardening, service suppression, and scheduled task enforcement. Creates a self-healing hourly maintenance task when Foxit is detected. |
| **MACE** | Microsoft Application Cleanse & Eradication | Completely removes OneDrive, New Outlook, Office/M365, Microsoft Project, and Microsoft Teams. Clears registry keys, cached credentials, profile data, and temp folders. Repairs OneDrive shell folder redirects and path pollution. |
| **ODD** | Output Device Diagnostic | Inventories all physical, USB, and Bluetooth audio devices with health status and driver versions across input and output categories. Highlights the default device and reports Audio Service state, sample rate, bit depth, and exclusive mode. |
| **ORCA** | Outlook Repair & Configuration Assistant | Resets broken Outlook installations for one or all user profiles. Supports New Outlook, Classic Outlook, or both. Clears cached data, registry hives, and authentication tokens. Optionally removes OST files and purges third-party extensions. |
| **PRINT** | Printer Response & Interface Network Tool | Printer management utility with a multi-threaded network scanner that discovers printers on the local subnet via port scan, identifies models through HTTP scraping, and automates driver matching from the local store. Supports manual IP and UNC paths. |
| **PRUNE** | Profile Removal Utility for Neglected Entries | Scans local Windows user profiles sorted by last-used date, flagging stale, orphaned, and disabled accounts for safe removal. Includes investigation mode to identify services and scheduled tasks keeping profile hives mounted. |
| **SHADE** | System Hardening Against Data Exposure | Comprehensive Windows 10/11 privacy hardening for security-conscious environments. Disables location tracking, telemetry services, advertising profiling, camera and microphone access, activity history, clipboard logging, feedback collection, Delivery Optimization, and network-level phone-home behavior. Safe for MSP deployment on Pro, Business, and Enterprise SKUs. A reboot is recommended after running. |
| **STARE** | Scheduled Task Administration & Routine Executor | Interactive terminal wizard for Windows Scheduled Task management. Supports Daily, Weekly, Monthly, and Startup triggers. Run a command or browse the filesystem to select a script, with a progressive UI that keeps selections in view at every step. |
| **TICK** | Trigger Immediate Clock Kickstart | Resets and resyncs the Windows Time service against a chosen NTP peer. Reports before/after timestamps with a plain-English summary of sync accuracy, stratum, and clock health. Domain-aware — detects DCs and defaults accordingly. |
| **URT** | Universal Rename Tool | Renames local or domain-joined computers from an admin shell with no GUI popups. Collects domain credentials inline if needed, preserves AD trust relationships, and offers an optional immediate reboot on completion. |
| **VITALS** | Visual Interface for Technical Asset & Logistics Summary | Full hardware and network snapshot: make/model/serial, CPU/RAM/GPU specs, disk health, BitLocker status and recovery keys, network configuration, battery wear, TPM/Secure Boot status, domain membership, and local admin members. |
| **WURSA** | Windows Update, Repair & System Alignment | Enforces all essential and optional OS patches, OEM driver updates, and third-party app upgrades via Chocolatey. Skips apps currently in use. Self-installs Chocolatey if not present. Includes a reliable unattended Windows feature upgrade with ESC-to-cancel. |

---

## Usage

All scripts require an elevated PowerShell session. Right-click → **Run with PowerShell** or launch from an admin terminal:

```powershell
.\VITALS.ps1
```

Most scripts are fully interactive. Several support parameters for unattended/RMM use — check the `.SYNOPSIS` block at the top of each file for available parameters and exit codes.

**Compatibility:** PowerShell 5.1 · PowerShell 7 · Kaseya LiveConnect

---

## Notes

- Scripts are standalone with no module dependencies unless noted in the header
- All scripts are tested on Windows 10 and Windows 11
- Use at your own risk in production — always test in a lab environment first
