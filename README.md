# Profile Rebuild Tool

**Automated fix for Excel/Outlook freezing and high CPU caused by a corrupted Windows user profile.**

## Why this exists

As a senior network engineer, I've been performing some version of this profile rebuild by hand for about five years, going after a handful of Windows 10/11 issues that all trace back to the same root cause: a corrupted user profile. I first learned the underlying technique nine years ago from [NLB Solutions](https://www.youtube.com/@NLBSolutions)' YouTube video, ["How to re-create a corrupted profile in Windows 10 (Step by Step guide)"](https://www.youtube.com/watch?v=RvUgAmT4kNQ), and the process itself was also heavily inspired by [User Profile Wizard from ForensiT Software Ltd.](https://www.forensit.com/), a commercial tool built around the same core idea of migrating/rebuilding a Windows profile while preserving the user's data and settings.

Originally, I reached for this fix for a narrower set of problems — Microsoft Office apps failing to activate, and Outlook classic getting stuck partway through sign-in when setting up a new profile. Over time, and especially recently, the same manual process turned out to resolve a much wider range of symptoms tied to profile corruption: Excel throttling the CPU during completely normal use, and various line-of-business (LOB) application functionality failures and glitches. The Excel CPU issue ended up being the most heavily documented and reproducible case (see [Root Cause](#root-cause) below), which is what finally pushed me to formalize the manual steps into this tool rather than keep repeating them by hand every time.

Doing this by hand is tedious and easy to get wrong: you're renaming profile folders, deleting registry keys, and manually tracking down everything OneDrive doesn't sync (bookmarks, signatures, ribbon customizations, mapped drives, printers) so the user doesn't lose it. One missed step and you've either broken the rebuild or the user comes back annoyed that their taskbar and Chrome bookmarks are gone.

This tool automates that whole workflow into a two-phase, point-and-click process, specifically so it's **safe to hand to a junior engineer who has never edited the registry before.** It's meant as a last resort — reach for it when a problem has been isolated to one user's profile (i.e., it doesn't reproduce on another profile on the same machine), and every less-invasive fix has already been ruled out.

> **Scope note:** This tool touches the local machine only. It runs the rebuild locally, on the affected machine, from a separate local admin account — not remotely, and not from the affected user's own session.

---

## Table of Contents

- [When to Use This](#when-to-use-this)
- [Symptoms](#symptoms)
- [Root Cause](#root-cause)
- [Diagnostic Checklist (Reference)](#diagnostic-checklist-reference)
- [Resolution](#resolution)
  - [Method 1 (Recommended): Profile Rebuild Tool](#method-1-recommended-profile-rebuild-tool)
  - [Known Issues / Limitations](#known-issues--limitations)
  - [Method 2 (Fallback): Manual Profile Rebuild](#method-2-fallback-manual-profile-rebuild)
- [Appendix: Technical Reference](#appendix-technical-reference)
- [Testing Checklist Before Using on a Live User](#testing-checklist-before-using-on-a-live-user)

---

## When to Use This

This tool was built around one specific, heavily-documented case — Excel freezing / high CPU that **survives** an Office repair, add-in troubleshooting, and antivirus exclusion testing, especially when the issue does **not** occur in Safe Mode (`excel /safe`) — but the underlying rebuild has been used successfully for a broader set of profile-corruption symptoms over the years, including:

- Microsoft Office apps failing to activate
- Outlook classic getting stuck or failing to complete sign-in when setting up a new profile
- Excel throttling CPU usage during otherwise normal operation (the case documented in detail below)
- Line-of-business (LOB) application functionality failures and glitches that trace back to a corrupted profile

In all of these cases, the common thread is the same: the problem is isolated to one user's profile — it doesn't reproduce on another profile on the same machine — and standard troubleshooting (repair, reinstall, cache clearing, add-in/AV exclusions) hasn't resolved it.

*Category: Desktop Support – Microsoft Office / General Windows Profile Issues · Applies to: Windows 10/11, Microsoft 365*

## Symptoms

The Excel CPU case is the most thoroughly documented and reproducible, so its symptoms are detailed below. The same root cause (profile corruption) has also produced the Office activation failures, Outlook classic sign-in failures, and LOB application glitches described above — those cases just haven't been captured with the same level of diagnostic detail as the Excel one.

**Excel-specific symptoms:**

- Excel shows "Not Responding" during normal operations (opening files, editing, calculating)
- Excel CPU usage rises to 75–96%, with total system CPU reaching up to 99%
- Excel is normal at idle (roughly 17–18% CPU, ~1.2 GB RAM) but spikes immediately on file changes
- The same workbook/functions work fine on a different machine
- Excel runs normally in Safe Mode (`excel /safe`) — CPU stays near 0%
- May be accompanied by unrelated symptoms on the same profile, such as Outlook classic failing to sign in

## Root Cause

*The following diagnostic detail is specific to the Excel CPU case, since that's the instance that was captured and analyzed in depth. The Office activation, Outlook sign-in, and LOB application cases share the same underlying cause — profile-scoped corruption — without an equivalent Procmon trace on file.*

A Process Monitor (Procmon) capture during a live freeze showed Excel issuing roughly 12,000 registry operations per second continuously, with 71% of all captured activity (over 100,000 operations in an 11.5-second window) consisting of a single repeating loop against one COM interface:

```
HKCU\Software\Classes\Interface\{00020846-0000-0000-C000-000000000046}\ProxyStubClsid32
```

This registry path consistently returned `NAME NOT FOUND`, causing Excel to repeatedly fall back to the HKLM Click-to-Run virtualized registry and HKCR before retrying the same lookup. Because the missing registration lives under the user's own registry hive (`HKCU`), this is **user-profile-scoped corruption**, not a machine-wide, add-in, antivirus, or workbook-specific issue — which explains why none of the earlier remediation attempts had any effect.

## Diagnostic Checklist (Reference)

The following steps were performed, in order, before the root cause was identified. Use this as a reference for what's already been ruled out on similar tickets, or as a troubleshooting sequence if a profile rebuild isn't yet warranted.

| Step Taken | Result | Conclusion |
|---|---|---|
| Confirmed Excel process was the source of high CPU via Task Manager | Confirmed — Excel itself, 75–96% CPU | Ruled in: Excel process, not another app |
| Uninstalled and reinstalled Microsoft Office | No change | Not an install corruption |
| Ran Microsoft Office Online Repair | Completed successfully, issue persisted | Not fixed by standard repair |
| Launched Excel in Safe Mode (`excel /safe`) | CPU stayed near 0%, issue did not occur | Something outside Safe Mode's disabled set is the cause |
| Disabled all COM add-ins (Acrobat, LogiOptions, Data Streamer, Power Map, Power Pivot) | No change | Not a COM add-in |
| Reviewed standard Excel add-ins (Analysis ToolPak, Solver, etc.) | None were enabled | Not a standard add-in |
| Tested with a workbook containing no external/SharePoint links | Same freeze occurred | Not the specific workbook or external link |
| Reset Excel-specific registry key (`HKCU\...\Office\16.0\Excel`) | No change | Not Excel-app-level registry settings |
| Cleared Office document cache (OfficeFileCache) and Excel diagnostics | No change | Not a stale cache issue |
| Checked Microsoft Defender real-time protection status | Confirmed not running | Not Defender |
| Disabled Malwarebytes real-time protection | No change | Not antivirus/real-time scanning |
| Installed all available firmware updates and rebooted | No change | Not a firmware/driver gap |
| Captured and analyzed a Procmon trace during a live freeze | Found ~12,000 registry ops/sec, 71% of all activity looping on one COM interface `ProxyStubClsid32` lookup returning `NAME NOT FOUND` | **Root cause identified:** user-profile-scoped COM registration corruption (`HKCU\Software\Classes`) |

## Resolution

There are two ways to perform the profile rebuild:

- **Method 1 (recommended)** — the Profile Rebuild Tool. Faster, and automates backup/restore of everything OneDrive doesn't sync.
- **Method 2 (fallback)** — the original manual process. Kept here in case the tool is unavailable, or a step needs to be done by hand for troubleshooting.

### Method 1 (Recommended): Profile Rebuild Tool

A PowerShell + WPF GUI tool that automates the full rebuild workflow, including backing up settings that OneDrive doesn't sync: browser bookmarks, Office ribbon/Quick Access Toolbar customizations, Power BI Desktop UI settings, default app associations, taskbar pins, signatures, custom dictionary, mapped drives, and printer connections.

#### Requirements

- Must be run **locally** on the affected machine (no remote-machine support in this version)
- Must be run from a **separate local admin account** — NOT from the affected user's own profile, since Phase 1 renames that profile's folder and unloads its registry hive, which can't be done while it's the active session
- The affected user must be **fully logged off** before Phase 1 completes the rename/registry-delete steps
- Close the tool after Phase 1 completes and re-launch it fresh for Phase 2 — see [Known Issues](#known-issues--limitations) for why

#### Running the Tool

```powershell
powershell.exe -ExecutionPolicy Bypass -File .\ProfileRebuildTool.ps1
```

The tool self-elevates via UAC if not already run as admin.

#### Phase 1 — Prepare

1. Open the **"1. Prepare Profile"** tab
2. Select the affected user from the dropdown (click **Refresh** if they don't appear)
3. Confirm or change the backup destination folder (defaults to `C:\ProfileRebuildTool\Backups`)
4. Click **Run Prep** — if the tool detects the user is still logged in, it will warn and ask for confirmation before continuing
5. The tool backs up bookmarks, Office ribbon/QAT files, Power BI Desktop settings, default app associations, taskbar pins, signatures/dictionary/templates, mapped drives, and printer connections, then renames the profile folder (e.g. `C:\Users\jsmith` → `C:\Users\jsmith.corrupt_TIMESTAMP`) and deletes the corresponding `ProfileList` registry key
6. Note the backup folder path shown in the completion dialog — you'll need it for Phase 2
7. When prompted, choose to close the tool now (recommended) rather than leaving it open

#### Phase 2 — Restore

1. Have the user log back in — Windows will generate a brand-new profile automatically
2. Re-launch the tool fresh (a new process, not a window left open from Phase 1)
3. Switch to the **"2. Restore Settings"** tab
4. Select the same user from the dropdown
5. Browse to the backup folder created in Phase 1
6. Click **Run Restore** — bookmarks, Office ribbon/QAT files, Power BI settings, default apps, taskbar pins, and other backed-up items are copied into the new profile
7. For Firefox specifically: since Firefox creates a new randomly-named profile folder on first launch, the tool can't restore into it automatically. It stages the backed-up `places.sqlite` to the user's Desktop and logs a manual copy step instead of guessing a folder name
8. Close Office apps and Power BI Desktop before restoring their settings files, and have the user re-open them afterward to confirm ribbon/QAT customizations came back

> Screenshots of both phases (Prepare and Restore, each showing a completed run with its log output and completion dialog) have been omitted from this README. Reach out if you'd like the redacted versions added back under `docs/images/`.

### Known Issues / Limitations

**Taskbar pin restoration — confirmed broken for most pinned apps.**
Testing confirmed this isn't limited to Chrome — Outlook classic, Word, Zoom, and Teams all show the same behavior. Restored taskbar pins appear as blank placeholder icons rather than working shortcuts. This is expected with the current registry-blob restore method, not a bug specific to one app. The user will need to open each affected app once and manually re-pin it to the taskbar afterward. Both the documentation and the restore completion log call this out so it isn't mistaken for a failure.

**Running Phase 1 and Phase 2 in the same open window can fail.**
If the registry hive doesn't fully release after Phase 1 (a known PowerShell registry-provider quirk — see [The `reg unload` Quirk](#the-reg-unload-quirk)), Phase 2 can fail to mount that hive in the same session. The tool prompts to close itself after Phase 1 completes; always relaunch fresh for Phase 2 rather than switching tabs in a window left open since Phase 1.

Other limitations to be aware of:

- **Browser bookmark support:** Chrome, Edge, Brave (all Chromium-based, same JSON format) and Firefox (SQLite-based, manual-assist restore only)
- **Office ribbon/QAT support:** Word, Excel, PowerPoint, Outlook (all window types), and OneNote, via `.officeUI`/`.officeSL` files
- **Power BI Desktop** UI settings are restored via a direct copy of `User.zip` — close Power BI Desktop fully before restoring
- **Default app associations and taskbar pins** are restored on a best-effort basis — always spot-check with the user after restore
- **AD/domain accounts:** on-prem and hybrid-joined domain accounts should already work, since their SIDs use the same stable format as local accounts. Entra ID (Azure AD)-only joined devices use a different local SID format and are included but less tested — verify on a test machine in that environment before relying on it
- **No dry-run mode yet** — test on a throwaway local account before running against a live user if you haven't used the tool before

### Method 2 (Fallback): Manual Profile Rebuild

Use this method if the tool is unavailable, or if you need to perform/troubleshoot an individual step by hand.

1. Close all Office applications and sign the affected user out of their account completely
2. Rename the existing local profile folder (`C:\Users\<username>`) to invalidate it, preserving the original data for rollback/reference (e.g. append `.old_TIMESTAMP`)
3. Navigate to `HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList` and locate the subkey matching the user's SID
4. Export that key first as a backup (right-click → Export)
5. Delete the registry key (SID) associated with the old, renamed profile — this forces Windows to rebuild a fresh profile on next logon rather than reusing the stale reference
6. Log the user back in — Windows will generate a brand-new, clean user profile
7. Manually retrieve and restore: Chrome/Edge/other browser bookmarks, default app associations, taskbar layout, and any OneDrive-excluded files (signatures, custom dictionary, templates) from the renamed old profile folder
8. Reconfigure default apps as needed
9. Relaunch Excel (and/or Outlook) to confirm normal functionality with no CPU spike

---

## Appendix: Technical Reference

This section documents how the tool itself works internally. It's intended for anyone maintaining, extending, or troubleshooting the PowerShell script — not required reading to simply run it.

### Architecture Overview

The tool is a single `.ps1` file with no external dependencies beyond what ships with Windows. The GUI is built with WPF, defined as an in-line XAML string and parsed at runtime via `System.Windows.Markup.XamlReader` — there's no compiled `.exe`, XAML designer, or code-behind class involved.

### Core Technique: Hosting WPF Inside a Plain `.ps1`

```powershell
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = @'
<Window xmlns="..." xmlns:x="...">
...
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)
```

Any control given an `x:Name` in XAML (e.g. `<Button x:Name="btnRunPrep">`) can be retrieved with `$window.FindName("btnRunPrep")` and behaves like a normal .NET object from there.

### Core Technique: Event Wiring

PowerShell exposes WPF's event model as an `.Add_EventName()` method taking a scriptblock:

```powershell
$btnRunPrep.Add_Click({
    Invoke-ProfilePrep -Profile $profile -BackupRoot $txtPrepDest.Text
})
```

Scriptblocks close over variables in their parent scope, so handlers can freely reference controls and state defined earlier in the script without explicit parameter passing.

### Core Technique: Keeping the UI Responsive Without a Background Thread

All operations run synchronously on the UI thread rather than on a background runspace, to keep the script simple. After each log write, the script pumps the Windows message queue so the log textbox updates live:

```powershell
[System.Windows.Forms.Application]::DoEvents()
```

This is a known simplification, not a fully robust pattern — it can technically let a user click a button mid-operation. For a tool run deliberately, step by step, by one admin at a time, this tradeoff is acceptable.

### Core Technique: Self-Elevation

```powershell
$principal = New-Object Security.Principal.WindowsPrincipal(
    [Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    exit
}
```

`-Verb RunAs` triggers the UAC prompt. `$PSCommandPath` relaunches the same script elevated, then exits the original non-elevated instance.

### Core Technique: Why Registry Hive Mounting Works

This is the trickiest part conceptually. A logged-in user's `HKEY_CURRENT_USER` is really their `NTUSER.DAT` file, mounted by Windows at `HKEY_USERS\<their SID>`. If they're logged off, that file sits unmounted on disk — but it can be mounted manually with `reg.exe load HKU\<any-name> path\to\NTUSER.DAT`.

Critically, **local account SIDs do not change when a profile is rebuilt** — the SID belongs to the account, not the profile folder. Deleting the `ProfileList` entry and renaming the folder just tells Windows "this SID has no valid profile anymore, build a new one"; the account and its SID persist.

Because of that, the tool always mounts a hive using the account's real SID as the mount name (`reg.exe load HKU\<SID> ...`) rather than an arbitrary temp name. A `.reg` file exported during Phase 1 bakes in the literal path `HKEY_USERS\<SID>\...`, and since the SID is still valid after the rebuild, that same `.reg` file can be re-imported during Phase 2 with no path rewriting needed.

### The `reg unload` Quirk

```powershell
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Start-Sleep -Milliseconds 300
reg.exe unload $HiveInfo.RegExePath
```

PowerShell's `Registry::` provider can leave file handles open on a loaded hive even after the script is "done" with it, which makes `reg unload` fail with access denied. Forcing garbage collection before unloading is a widely used workaround — reliable in practice, though not a strict guarantee.

### Office Ribbon/QAT + Power BI Desktop Settings

Word, Excel, PowerPoint, Outlook, and OneNote store ribbon and Quick Access Toolbar customizations as `.officeUI` (classic ribbon) and `.officeSL` (simplified ribbon) files under `AppData\Local\Microsoft\Office`. Outlook creates several of these — one per window type (`olkexplorer.officeUI` for the main window, `olkmailread.officeUI` for a reading window, etc.) — so the tool backs up every matching file in that folder rather than assuming one fixed filename per app.

Power BI Desktop stores its UI/layout settings differently: inside a zip file (`User.zip`) containing `UserInterface\Settings.xml`. Since it's already a zip, the tool just copies that one file whole rather than unzipping and re-zipping it.

Both apps should be fully closed before restoring their settings files, since an open app can hold the file locked or overwrite the restored version with its own in-memory state on exit.

### AD / Domain Account Support

The SID-stability principle this tool relies on (see [Why Registry Hive Mounting Works](#core-technique-why-registry-hive-mounting-works) above) holds for on-prem and hybrid-joined domain accounts as well as local accounts: a domain account's SID uses the same `S-1-5-21-<domain>-<RID>` structure, and it belongs to the account object in Active Directory, not to any one machine's local profile. Rebuilding the local profile on this machine does not change that SID, so the same Phase 1 / Phase 2 flow applies without modification.

Devices that are Entra ID (Azure AD) joined **only** (not hybrid) can assign local account SIDs in a different format (`S-1-12-1-...`). The profile-detection logic includes this pattern, but it hasn't been as thoroughly tested as the `S-1-5-21-` case — verify on a test Entra-only-joined machine before relying on it in that environment.

### Design Decision: Close and Relaunch Between Phases

Earlier testing found that trying to run Phase 2 in the same window left open since Phase 1 could fail. The cause traces back to the `reg unload` quirk described above: if Phase 1's hive unload doesn't fully succeed, the hive stays locked for the life of that PowerShell process, not just until the next button click. Since process exit is the one thing that reliably releases any handle PowerShell's registry provider is holding, the tool now treats closing and relaunching between phases as the expected flow — it prompts to close itself right after Phase 1 completes — rather than trying to engineer around the underlying constraint.

### Browser Support Details

Chrome, Edge, and Brave are all Chromium-based and share the same plain-JSON `Bookmarks` file format, just under different folder paths — all handled by one generalized function pair (`Backup-BrowserBookmarks` / `Restore-BrowserBookmarks`). All profile folders under `User Data` (`Default`, `Profile 1`, `Profile 2`, ...) are backed up, not just `Default`.

Firefox is architecturally different: bookmarks live in a SQLite database (`places.sqlite`) inside a randomly-named profile folder, located by reading `profiles.ini`. Because a brand-new Firefox profile folder with a new random name doesn't exist until Firefox is launched once on the rebuilt profile, restore for Firefox is manual-assist only — the tool stages the backed-up `places.sqlite` to the user's Desktop and logs the manual copy step rather than guessing a folder name that doesn't exist yet.

Adding another Chromium-based browser (Opera, Vivaldi, etc.) is a one-line change — add an entry to the `$script:ChromiumBrowsers` array with its `User Data` relative path.

### Taskbar Pin Testing Matrix — Results

The `Taskband` registry value stores pinned items as serialized shell link data. Testing confirmed the failure mode (icon present, shortcut broken, shows as a blank placeholder) is not limited to Chrome or even to Chromium browsers — it also affects Outlook classic, Word, Zoom, and Teams. This points to the `Taskband` registry-blob restore approach itself being unreliable in general, rather than any one app's pinning mechanism.

Results so far:

- **Chrome, Outlook classic, Word, Zoom, Teams** — CONFIRMED BROKEN: restored pin shows as a blank placeholder icon, requires manual re-pin after first opening the app
- **Store/UWP apps** (Outlook new, Calculator, Photos) — not yet tested
- **Traditional Win32 apps other than Word** (Excel, Notepad, etc.) — not yet tested, but given Word's result, assume the same until proven otherwise

Given the breadth of confirmed failures, the current recommendation is to treat taskbar restoration as informational at best — the documentation now sets the expectation up front that pinned icons will need to be manually re-pinned, rather than presenting the registry restore as reliable. A future version could replace the `Taskband` blob restore with re-pinning apps individually from a plain recorded list instead, which would sidestep this failure mode entirely.

### What's Deliberately Simple in This Version

- **No dry-run mode** — consider adding a checkbox that logs intended actions without performing them, for safely training new techs on the tool
- **No multi-machine/remote support** — only touches the local machine; `Invoke-Command -ComputerName` is the natural next step, though hive loading gets more complex remotely
- **Extra backed-up items list is hardcoded** in `Backup-UserData` (`$extraFilePaths`) — add more relative paths there as more OneDrive-excluded items are discovered
- **No confirmation diff before restore** — running Restore twice for the same user simply overwrites; add a warning dialog if this tool is extended for broader team use

## Testing Checklist Before Using on a Live User

1. Run Phase 1 on a test/throwaway local account first, not a live user
2. Confirm the backup folder actually contains bookmarks/reg files as expected
3. Log the test user off, back on, confirm Windows built a fresh profile
4. Run Phase 2, confirm bookmarks and default apps came back
5. Only then run it against an actual affected user
