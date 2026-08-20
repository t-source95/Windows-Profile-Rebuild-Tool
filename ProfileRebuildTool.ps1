#Requires -Version 5.1
<#
.SYNOPSIS
    Profile Rebuild Tool - GUI-driven local Windows profile rebuild + settings restore.

.DESCRIPTION
    Two-phase workflow for fixing corrupted local user profiles (Excel/Outlook loop bugs,
    unreadable NTUSER.DAT, etc.) without losing personal customizations that OneDrive
    doesn't sync.

    PHASE 1 - PREPARE (run while the user is still logged in, before you sign them out):
      - Backs up Chrome/Edge bookmarks, default app associations, taskbar pinned items,
        signatures, custom dictionary, mapped drives, and printer connections
      - Renames the profile folder (e.g. C:\Users\ljohnson -> C:\Users\ljohnson.corrupt_TIMESTAMP)
      - Backs up and deletes the corresponding key under
        HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList
      - This forces Windows to build a brand-new profile on next logon

    PHASE 2 - RESTORE (run after the user has logged back in and Windows has created
    a fresh profile):
      - Copies bookmarks back into the new browser profile
      - Re-imports default app associations
      - Re-imports taskbar pinned items
      - Restores signatures/dictionary/misc files

.NOTES
    Requires local admin rights. Self-elevates if not already running as admin.
    Designed for local (non-domain-SID-changing) accounts, since the account SID stays
    the same across a profile rebuild - only the profile folder and ProfileList entry
    are new. That stability is what makes Phase 2 possible.
#>

# ============================================================================
# SELF-ELEVATION
# If this script isn't running as admin, relaunch itself elevated and exit
# the current (non-elevated) instance.
# ============================================================================
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`"" `
        -Verb RunAs
    exit
}

# ============================================================================
# ASSEMBLIES
# PresentationFramework/Core/WindowsBase = WPF. System.Windows.Forms is only
# used here for the FolderBrowserDialog, since WPF has no built-in one.
# ============================================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Xaml

# ============================================================================
# GLOBAL STATE
# ============================================================================
$script:Profiles      = @()
$script:LogBox         = $null
$script:LogFilePath    = $null
$ToolRoot              = "C:\ProfileRebuildTool"
$DefaultBackupRoot     = Join-Path $ToolRoot "Backups"
New-Item -ItemType Directory -Path $ToolRoot -Force        | Out-Null
New-Item -ItemType Directory -Path $DefaultBackupRoot -Force | Out-Null

# ============================================================================
# LOGGING
# Writes to both the GUI textbox and a log file, and pumps the WPF message
# queue so the log appears live even though everything runs synchronously
# on the UI thread (no background runspace in this v1).
# ============================================================================
function Write-Log {
    param([string]$Message)
    $line = "[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $Message
    if ($script:LogBox) {
        $script:LogBox.AppendText("$line`r`n")
        $script:LogBox.ScrollToEnd()
    }
    if ($script:LogFilePath) {
        Add-Content -Path $script:LogFilePath -Value $line
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# ============================================================================
# PROFILE ENUMERATION
# Reads HKLM\...\ProfileList and resolves each SID to a username.
# Only real user account SIDs, not system/service accounts:
#   S-1-5-21-* covers local accounts AND on-prem/hybrid AD domain accounts
#     (a domain account's SID uses the same S-1-5-21-<domain>-<RID> structure,
#     and it doesn't change when this machine's local profile is rebuilt -
#     the SID belongs to the account object in AD, not to this machine).
#   S-1-12-1-* covers accounts on devices that are Entra ID (Azure AD) joined
#     ONLY (not hybrid) - these use a different local SID scheme. Included
#     here but less tested; verify on an Entra-only-joined test machine
#     before relying on it in that environment.
function Get-LocalProfiles {
    $base = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList'
    Get-ChildItem $base | ForEach-Object {
        $sid = $_.PSChildName
        if ($sid -match '^S-1-5-21-' -or $sid -match '^S-1-12-1-') {
            $props = Get-ItemProperty -Path $_.PSPath -ErrorAction SilentlyContinue
            $imagePath = $props.ProfileImagePath
            $username = "Unknown ($sid)"
            try {
                $username = (New-Object System.Security.Principal.SecurityIdentifier($sid)).
                            Translate([System.Security.Principal.NTAccount]).Value
            } catch { }
            [PSCustomObject]@{
                SID          = $sid
                Username     = $username
                ProfilePath  = $imagePath
                RegistryKey  = $_.PSPath
                Exists       = (Test-Path $imagePath)
            }
        }
    }
}

function Update-ProfileComboBox {
    param($ComboBox)
    $script:Profiles = @(Get-LocalProfiles | Sort-Object Username)
    $ComboBox.Items.Clear()
    foreach ($p in $script:Profiles) {
        $tag = if ($p.Exists) { "" } else { "  [MISSING FOLDER]" }
        [void]$ComboBox.Items.Add("$($p.Username)   [$($p.SID)]$tag")
    }
}

function Get-SelectedProfile {
    param($ComboBox)
    if ($ComboBox.SelectedIndex -lt 0) { return $null }
    return $script:Profiles[$ComboBox.SelectedIndex]
}

# ============================================================================
# REGISTRY HIVE MOUNTING
# The tricky part: to back up or restore a user's HKCU-scoped settings
# (taskbar pins, mapped drives, printer connections), that user's registry
# hive needs to be accessible. If they're logged in, Windows already mounts
# it live at HKEY_USERS\<SID>. If they're logged off, we load their
# NTUSER.DAT manually with reg.exe and mount it under that SAME SID name.
#
# Using the real SID as the mount name (rather than an arbitrary temp name)
# is what lets exported .reg files stay valid: a .reg file bakes in the full
# path text (e.g. "HKEY_USERS\S-1-5-21-...\Software\..."), and since a local
# account's SID never changes across a profile rebuild, a .reg exported
# during Phase 1 can be re-imported during Phase 2 without any path rewriting.
# ============================================================================
function Mount-UserHive {
    param([string]$SID, [string]$ProfilePath)

    $liveHivePath = "Registry::HKEY_USERS\$SID"
    if (Test-Path $liveHivePath) {
        return [PSCustomObject]@{ Path = $liveHivePath; Loaded = $false; RegExePath = "HKU\$SID" }
    }

    $ntuser = Join-Path $ProfilePath "NTUSER.DAT"
    if (Test-Path $ntuser) {
        $result = & reg.exe load "HKU\$SID" "$ntuser" 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "WARNING: could not load hive for $SID - $result"
            return $null
        }
        Start-Sleep -Milliseconds 300
        return [PSCustomObject]@{ Path = $liveHivePath; Loaded = $true; RegExePath = "HKU\$SID" }
    }

    Write-Log "WARNING: no NTUSER.DAT found at $ntuser - skipping registry-based backup items"
    return $null
}

function Dismount-UserHive {
    param($HiveInfo)
    if ($null -ne $HiveInfo -and $HiveInfo.Loaded) {
        # PowerShell's registry provider can leave handles open on a hive
        # even after you're "done" with it, which blocks reg unload.
        # Forcing garbage collection releases those handles reliably enough
        # for this to work in practice - a known quirk, not a guarantee.
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        Start-Sleep -Milliseconds 300
        & reg.exe unload $HiveInfo.RegExePath 2>&1 | Out-Null
    }
}

# ============================================================================
# ACTIVE SESSION CHECK
# Renaming a profile folder or unloading its hive fails if the user is still
# logged in and files/registry are locked. `query user` tells us if they are.
# ============================================================================
function Test-UserLoggedIn {
    param([string]$Username)
    $result = & query.exe user $Username 2>&1
    return ($LASTEXITCODE -eq 0)
}

# ============================================================================
# BROWSER BOOKMARKS - Chromium family (Chrome, Edge, Brave) + Firefox
#
# Chrome/Edge/Brave share the same Chromium bookmark format: a plain JSON
# file at "User Data\<profile>\Bookmarks". Firefox is architecturally
# different - it stores bookmarks in a SQLite database (places.sqlite)
# inside a randomly-named profile folder, so we read profiles.ini first to
# find the right one instead of assuming a fixed path.
# ============================================================================

# Chromium-based browsers: relative path from the user's profile root to the
# "User Data" folder. Add more browsers here (e.g. Opera, Vivaldi) by adding
# another entry - they all use this same structure.
$script:ChromiumBrowsers = @(
    @{ Name = "Chrome"; RelPath = "AppData\Local\Google\Chrome\User Data" }
    @{ Name = "Edge";   RelPath = "AppData\Local\Microsoft\Edge\User Data" }
    @{ Name = "Brave";  RelPath = "AppData\Local\BraveSoftware\Brave-Browser\User Data" }
)

function Get-ChromiumProfileFolders {
    # A Chromium "User Data" folder can contain multiple profiles (Default,
    # "Profile 1", "Profile 2", ...). We back up bookmarks from all of them
    # rather than assuming only Default is in use.
    param([string]$UserDataPath)
    if (-not (Test-Path $UserDataPath)) { return @() }
    Get-ChildItem $UserDataPath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -eq "Default" -or $_.Name -match '^Profile \d+$' } |
        Where-Object { Test-Path (Join-Path $_.FullName "Bookmarks") }
}

function Backup-BrowserBookmarks {
    param([string]$ProfilePath, [string]$DestFolder)

    $browserDest = Join-Path $DestFolder "Browsers"

    # --- Chromium family ---
    foreach ($browser in $script:ChromiumBrowsers) {
        $userDataPath = Join-Path $ProfilePath $browser.RelPath
        $profileFolders = Get-ChromiumProfileFolders -UserDataPath $userDataPath
        foreach ($pf in $profileFolders) {
            $bookmarksFile = Join-Path $pf.FullName "Bookmarks"
            $outDir = Join-Path $browserDest "$($browser.Name)\$($pf.Name)"
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            Copy-Item $bookmarksFile (Join-Path $outDir "Bookmarks") -Force
            Write-Log "Backed up $($browser.Name) bookmarks ($($pf.Name))"
        }
    }

    # --- Firefox ---
    $ffProfilesIni = Join-Path $ProfilePath "AppData\Roaming\Mozilla\Firefox\profiles.ini"
    if (Test-Path $ffProfilesIni) {
        $ffProfilesRoot = Join-Path $ProfilePath "AppData\Roaming\Mozilla\Firefox"
        $iniContent = Get-Content $ffProfilesIni -Raw
        # profiles.ini is a plain INI file; each [Profile#] section has a
        # Path= line. We don't need a full INI parser for this - a regex
        # over each section is enough.
        $sections = $iniContent -split '(?=\[Profile\d+\])' | Where-Object { $_ -match '^\[Profile\d+\]' }
        foreach ($section in $sections) {
            if ($section -match 'Path=(.+)') {
                $relProfilePath = $matches[1].Trim()
                $isRelative = $section -match 'IsRelative=1'
                $fullProfilePath = if ($isRelative) { Join-Path $ffProfilesRoot $relProfilePath } else { $relProfilePath }
                $placesDb = Join-Path $fullProfilePath "places.sqlite"
                if (Test-Path $placesDb) {
                    $profileName = Split-Path $fullProfilePath -Leaf
                    $outDir = Join-Path $browserDest "Firefox\$profileName"
                    New-Item -ItemType Directory -Path $outDir -Force | Out-Null
                    # Copy the whole places.sqlite (bookmarks + history) plus
                    # its write-ahead log files if present, since SQLite can
                    # keep pending writes there rather than in the main file.
                    Copy-Item $placesDb (Join-Path $outDir "places.sqlite") -Force
                    foreach ($ext in @("-wal", "-shm")) {
                        $sidecar = "$placesDb$ext"
                        if (Test-Path $sidecar) {
                            Copy-Item $sidecar (Join-Path $outDir "places.sqlite$ext") -Force
                        }
                    }
                    Write-Log "Backed up Firefox bookmarks ($profileName)"
                }
            }
        }
    }
}

function Restore-BrowserBookmarks {
    param([string]$ProfilePath, [string]$BackupFolder)

    $browserBackupRoot = Join-Path $BackupFolder "Browsers"
    if (-not (Test-Path $browserBackupRoot)) { return }

    # --- Chromium family: copy each backed-up profile folder back to the
    # matching relative location under the (new) user profile. Because this
    # is a fresh profile, the browser likely hasn't been launched yet, so
    # "Default" may not exist - create it if needed. ---
    foreach ($browser in $script:ChromiumBrowsers) {
        $browserBackupDir = Join-Path $browserBackupRoot $browser.Name
        if (-not (Test-Path $browserBackupDir)) { continue }
        $userDataPath = Join-Path $ProfilePath $browser.RelPath
        Get-ChildItem $browserBackupDir -Directory | ForEach-Object {
            $profileName = $_.Name
            $destProfileDir = Join-Path $userDataPath $profileName
            New-Item -ItemType Directory -Path $destProfileDir -Force | Out-Null
            Copy-Item (Join-Path $_.FullName "Bookmarks") (Join-Path $destProfileDir "Bookmarks") -Force
            Write-Log "Restored $($browser.Name) bookmarks ($profileName)"
        }
    }

    # --- Firefox: restoring is trickier because a *new* profile folder with
    # a *new* random name will be created the first time Firefox runs on the
    # rebuilt profile. We can't restore into it before it exists. Simplest
    # reliable approach: stage the backed-up places.sqlite in a known location
    # and log clear manual instructions, rather than guessing a profile name
    # that doesn't exist yet. ---
    $ffBackupDir = Join-Path $browserBackupRoot "Firefox"
    if (Test-Path $ffBackupDir) {
        $stagingDir = Join-Path $ProfilePath "Desktop\_RestoredFirefoxBookmarks"
        New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null
        Copy-Item (Join-Path $ffBackupDir "*") $stagingDir -Recurse -Force
        Write-Log "Firefox bookmarks staged to Desktop\_RestoredFirefoxBookmarks"
        Write-Log "MANUAL STEP: launch Firefox once to create its new profile, close it,"
        Write-Log "  then copy places.sqlite from the staged folder into:"
        Write-Log "  %APPDATA%\Mozilla\Firefox\Profiles\<new-profile-folder>\"
    }
}

# ============================================================================
# OFFICE RIBBON / QUICK ACCESS TOOLBAR + POWER BI DESKTOP UI SETTINGS
#
# Word, Excel, PowerPoint, Outlook, and OneNote store ribbon and Quick Access
# Toolbar customizations as .officeUI (classic ribbon) and .officeSL
# (simplified ribbon) files under AppData\Local\Microsoft\Office. Outlook in
# particular creates several of these - one per window type (e.g.
# olkexplorer.officeUI for the main window, olkmailread.officeUI for reading
# a message, etc.) - so all matching files in that folder are backed up
# rather than assuming one fixed filename per app.
#
# Power BI Desktop stores its UI/layout settings differently: inside a zip
# file (User.zip) containing UserInterface\Settings.xml. Since it's already
# a zip, the simplest reliable backup is just copying that one file whole -
# no need to unzip/rezip it ourselves.
# ============================================================================
function Backup-OfficeAndPowerBiSettings {
    param([string]$ProfilePath, [string]$DestFolder)

    # --- Office ribbon / QAT customization files ---
    $officeUiPath = Join-Path $ProfilePath "AppData\Local\Microsoft\Office"
    if (Test-Path $officeUiPath) {
        $uiFiles = Get-ChildItem $officeUiPath -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -in ".officeUI", ".officeSL" }
        if ($uiFiles) {
            $outDir = Join-Path $DestFolder "OfficeRibbonUI"
            New-Item -ItemType Directory -Path $outDir -Force | Out-Null
            foreach ($f in $uiFiles) {
                Copy-Item $f.FullName (Join-Path $outDir $f.Name) -Force
            }
            Write-Log "Backed up $($uiFiles.Count) Office ribbon/QAT customization file(s)"
        }
    }

    # --- Power BI Desktop UI settings ---
    $pbiUserZip = Join-Path $ProfilePath "AppData\Local\Microsoft\Power BI Desktop\User.zip"
    if (Test-Path $pbiUserZip) {
        $outDir = Join-Path $DestFolder "PowerBIDesktop"
        New-Item -ItemType Directory -Path $outDir -Force | Out-Null
        Copy-Item $pbiUserZip (Join-Path $outDir "User.zip") -Force
        Write-Log "Backed up Power BI Desktop UI settings (User.zip)"
    }
}

function Restore-OfficeAndPowerBiSettings {
    param([string]$ProfilePath, [string]$BackupFolder)

    # --- Office ribbon / QAT customization files ---
    $uiBackupDir = Join-Path $BackupFolder "OfficeRibbonUI"
    if (Test-Path $uiBackupDir) {
        $officeUiPath = Join-Path $ProfilePath "AppData\Local\Microsoft\Office"
        New-Item -ItemType Directory -Path $officeUiPath -Force | Out-Null
        Get-ChildItem $uiBackupDir -File | ForEach-Object {
            Copy-Item $_.FullName (Join-Path $officeUiPath $_.Name) -Force
        }
        Write-Log "Restored Office ribbon/QAT customization files (close Office apps first if any are open)"
    }

    # --- Power BI Desktop UI settings ---
    $pbiBackup = Join-Path $BackupFolder "PowerBIDesktop\User.zip"
    if (Test-Path $pbiBackup) {
        $pbiDestDir = Join-Path $ProfilePath "AppData\Local\Microsoft\Power BI Desktop"
        New-Item -ItemType Directory -Path $pbiDestDir -Force | Out-Null
        Copy-Item $pbiBackup (Join-Path $pbiDestDir "User.zip") -Force
        Write-Log "Restored Power BI Desktop UI settings (Power BI Desktop must be fully closed first)"
    }
}

# ============================================================================
# PHASE 1 STEP: BACKUP
# ============================================================================
function Backup-UserData {
    param($Profile, [string]$DestFolder)

    New-Item -ItemType Directory -Path $DestFolder -Force | Out-Null
    $profilePath = $Profile.ProfilePath
    $sid         = $Profile.SID

    # --- Browser bookmarks ---
    Backup-BrowserBookmarks -ProfilePath $profilePath -DestFolder $DestFolder

    # --- Office ribbon/QAT customizations + Power BI Desktop UI settings ---
    Backup-OfficeAndPowerBiSettings -ProfilePath $profilePath -DestFolder $DestFolder

    # --- Misc files OneDrive typically does NOT sync ---
    $extraFilePaths = @(
        "AppData\Roaming\Microsoft\Signatures",   # Outlook signatures
        "AppData\Roaming\Microsoft\UProof",        # custom dictionary
        "AppData\Roaming\Microsoft\Templates"      # Office templates
    )
    $extraDest = Join-Path $DestFolder "ExtraItems"
    foreach ($rel in $extraFilePaths) {
        $src = Join-Path $profilePath $rel
        if (Test-Path $src) {
            $dst = Join-Path $extraDest $rel
            New-Item -ItemType Directory -Path (Split-Path $dst -Parent) -Force | Out-Null
            Copy-Item $src $dst -Recurse -Force
            Write-Log "Backed up $rel"
        }
    }

    # --- Default app associations (system-wide DISM export, not per-user,
    #     but relevant per machine and cheap to capture each time) ---
    $assocFile = Join-Path $DestFolder "DefaultAppAssociations.xml"
    try {
        & Dism.exe /Online /Export-DefaultAppAssociations:"$assocFile" | Out-Null
        Write-Log "Exported default app associations"
    } catch {
        Write-Log "Could not export default app associations: $_"
    }

    # --- Registry-based per-user items: taskbar pins, mapped drives, printers ---
    $hive = Mount-UserHive -SID $sid -ProfilePath $profilePath
    if ($hive) {
        $regTargets = @(
            @{ Path = "Software\Microsoft\Windows\CurrentVersion\Explorer\Taskband"; File = "Taskband.reg" }
            @{ Path = "Network";                                                     File = "MappedDrives.reg" }
            @{ Path = "Printers\Connections";                                        File = "PrinterConnections.reg" }
        )
        foreach ($t in $regTargets) {
            $fullRegExePath = "$($hive.RegExePath)\$($t.Path)"
            $checkPath = "Registry::HKEY_USERS\$sid\$($t.Path)"
            if (Test-Path $checkPath) {
                $outFile = Join-Path $DestFolder $t.File
                & reg.exe export $fullRegExePath $outFile /y 2>&1 | Out-Null
                Write-Log "Backed up registry: $($t.Path)"
            }
        }
        Dismount-UserHive $hive
    }

    # Manifest for traceability
    $manifest = [PSCustomObject]@{
        Username    = $Profile.Username
        SID         = $sid
        SourcePath  = $profilePath
        BackupDate  = Get-Date -Format "o"
    }
    $manifest | ConvertTo-Json | Set-Content (Join-Path $DestFolder "manifest.json")
}

# ============================================================================
# PHASE 1: PREPARE
# ============================================================================
function Invoke-ProfilePrep {
    param($Profile, [string]$BackupRoot)

    if (Test-UserLoggedIn -Username ($Profile.Username -replace '.*\\','')) {
        $result = [System.Windows.MessageBox]::Show(
            "$($Profile.Username) still appears to be logged in. They must be fully logged off " +
            "before the profile folder can be renamed and the hive unloaded.`n`nContinue anyway?",
            "User still logged in", "YesNo", "Warning")
        if ($result -eq "No") { Write-Log "Prep cancelled - user still logged in."; return }
    }

    $timestamp  = Get-Date -Format "yyyyMMdd_HHmmss"
    $safeName   = ($Profile.Username -replace '[\\\/:*?"<>|]', '_')
    $destFolder = Join-Path $BackupRoot "${safeName}_$timestamp"
    $script:LogFilePath = Join-Path $destFolder "prep-log.txt"
    New-Item -ItemType Directory -Path $destFolder -Force | Out-Null

    Write-Log "=== PHASE 1: PREPARE - $($Profile.Username) ==="
    Write-Log "Backup destination: $destFolder"

    Backup-UserData -Profile $Profile -DestFolder $destFolder

    # Back up the ProfileList key itself before touching anything
    $regKeyPath = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($Profile.SID)"
    $regBackupFile = Join-Path $destFolder "ProfileListKey_backup.reg"
    & reg.exe export $regKeyPath $regBackupFile /y | Out-Null
    Write-Log "Backed up ProfileList registry key to $regBackupFile"

    # Rename the profile folder
    $oldPath = $Profile.ProfilePath
    $renamedLeaf = (Split-Path $oldPath -Leaf) + ".corrupt_$timestamp"
    try {
        Rename-Item -Path $oldPath -NewName $renamedLeaf -ErrorAction Stop
        Write-Log "Renamed profile folder to $renamedLeaf"
    } catch {
        Write-Log "ERROR renaming profile folder: $_"
        Write-Log "Make sure the user is fully logged off and no processes hold files open (check Task Manager > Users)."
        return
    }

    # Delete the ProfileList entry so Windows treats this as a new user
    try {
        Remove-Item -Path $Profile.RegistryKey -Recurse -Force -ErrorAction Stop
        Write-Log "Deleted ProfileList registry key for SID $($Profile.SID)"
    } catch {
        Write-Log "ERROR deleting ProfileList key: $_"
        return
    }

    Write-Log ""
    Write-Log "PREP COMPLETE."
    Write-Log "Next: have the user log back in to generate a fresh profile."
    Write-Log "Then re-launch this tool and use the Restore tab, pointing it at:"
    Write-Log "  $destFolder"
    Write-Log ""
    Write-Log "NOTE: closing this tool now (rather than leaving it open) is recommended -"
    Write-Log "it guarantees the user's registry hive is fully released before they log back in."

    $closeNow = [System.Windows.MessageBox]::Show(
        "Prep complete for $($Profile.Username).`n`nBackup saved to:`n$destFolder`n`n" +
        "Have the user log back in now, then re-launch this tool and use Restore with this backup folder.`n`n" +
        "Close this window now? (Recommended - guarantees the registry hive fully releases " +
        "before the user logs back in, rather than relying on this same session to release it later.)",
        "Prep Complete", "YesNo", "Information")
    if ($closeNow -eq "Yes") {
        $window.Close()
    }
}

# ============================================================================
# PHASE 2: RESTORE
# ============================================================================
function Invoke-ProfileRestore {
    param($Profile, [string]$BackupFolder)

    if (-not (Test-Path $BackupFolder)) {
        Write-Log "ERROR: backup folder not found: $BackupFolder"
        return
    }

    $script:LogFilePath = Join-Path $BackupFolder "restore-log.txt"
    Write-Log "=== PHASE 2: RESTORE - $($Profile.Username) ==="
    Write-Log "Restoring from: $BackupFolder"
    Write-Log "Target profile: $($Profile.ProfilePath)"

    $profilePath = $Profile.ProfilePath
    $sid         = $Profile.SID

    # --- Bookmarks (Chrome, Edge, Brave, Firefox) ---
    Restore-BrowserBookmarks -ProfilePath $profilePath -BackupFolder $BackupFolder

    # --- Office ribbon/QAT customizations + Power BI Desktop UI settings ---
    Restore-OfficeAndPowerBiSettings -ProfilePath $profilePath -BackupFolder $BackupFolder

    # --- Extra items ---
    $extraDest = Join-Path $BackupFolder "ExtraItems"
    if (Test-Path $extraDest) {
        Get-ChildItem $extraDest -Recurse -File | ForEach-Object {
            $relative = $_.FullName.Substring($extraDest.Length + 1)
            $target = Join-Path $profilePath $relative
            New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
            Copy-Item $_.FullName $target -Force
        }
        Write-Log "Restored extra items (signatures, dictionary, templates)"
    }

    # --- Default app associations ---
    $assocFile = Join-Path $BackupFolder "DefaultAppAssociations.xml"
    if (Test-Path $assocFile) {
        & Dism.exe /Online /Import-DefaultAppAssociations:"$assocFile" | Out-Null
        Write-Log "Imported default app associations"
    }

    # --- Registry items: taskbar, mapped drives, printers ---
    $hive = Mount-UserHive -SID $sid -ProfilePath $profilePath
    if ($hive) {
        foreach ($file in @("Taskband.reg", "MappedDrives.reg", "PrinterConnections.reg")) {
            $path = Join-Path $BackupFolder $file
            if (Test-Path $path) {
                & reg.exe import $path 2>&1 | Out-Null
                Write-Log "Restored registry: $file"
            }
        }
        Dismount-UserHive $hive
    }

    Write-Log ""
    Write-Log "RESTORE COMPLETE for $($Profile.Username)."
    Write-Log "Note: restored taskbar pins commonly show as blank placeholder icons - this is"
    Write-Log "  expected. Confirmed across Chrome, Outlook classic, Word, Zoom, and Teams."
    Write-Log "  The user will need to re-pin each app once after opening it."

    [System.Windows.MessageBox]::Show(
        "Restore complete for $($Profile.Username).",
        "Restore Complete", "OK", "Information") | Out-Null
}

# ============================================================================
# XAML - UI LAYOUT
# Loaded via XamlReader at runtime; no XAML compiler/designer needed.
# x:Name'd controls are retrieved with $window.FindName() after load.
# ============================================================================
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Profile Rebuild Tool" Height="680" Width="720"
        WindowStartupLocation="CenterScreen">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="*"/>
            <RowDefinition Height="220"/>
        </Grid.RowDefinitions>

        <TabControl Grid.Row="0">
            <TabItem Header="1. Prepare Profile">
                <StackPanel Margin="15">
                    <TextBlock Text="Select the affected user's profile:" FontWeight="Bold" Margin="0,0,0,5"/>
                    <DockPanel Margin="0,0,0,10">
                        <Button x:Name="btnRefreshPrep" Content="Refresh" Width="80" DockPanel.Dock="Right"/>
                        <ComboBox x:Name="cmbUsersPrep" Margin="0,0,10,0"/>
                    </DockPanel>

                    <TextBlock Text="Backup destination folder:" Margin="0,10,0,5"/>
                    <DockPanel Margin="0,0,0,15">
                        <Button x:Name="btnBrowsePrepDest" Content="Browse..." Width="80" DockPanel.Dock="Right"/>
                        <TextBox x:Name="txtPrepDest" Margin="0,0,10,0"/>
                    </DockPanel>

                    <TextBlock TextWrapping="Wrap" Foreground="#FF444444" Margin="0,0,0,15"
                        Text="This backs up bookmarks, Office ribbon/QAT customizations, Power BI Desktop UI settings, default apps, taskbar pins, signatures, mapped drives, and printers - then renames the profile folder and deletes its ProfileList registry entry. Make sure the user is fully logged off first. After Prep completes, close this tool and re-launch it fresh for Restore once the user has logged back in."/>

                    <Button x:Name="btnRunPrep" Content="Run Prep" Height="36" Background="#FFC0392B" Foreground="White" FontWeight="Bold"/>
                </StackPanel>
            </TabItem>

            <TabItem Header="2. Restore Settings">
                <StackPanel Margin="15">
                    <TextBlock Text="Select the user (after they've logged back in):" FontWeight="Bold" Margin="0,0,0,5"/>
                    <DockPanel Margin="0,0,0,10">
                        <Button x:Name="btnRefreshRestore" Content="Refresh" Width="80" DockPanel.Dock="Right"/>
                        <ComboBox x:Name="cmbUsersRestore" Margin="0,0,10,0"/>
                    </DockPanel>

                    <TextBlock Text="Backup folder to restore from:" Margin="0,10,0,5"/>
                    <DockPanel Margin="0,0,0,15">
                        <Button x:Name="btnBrowseRestoreSrc" Content="Browse..." Width="80" DockPanel.Dock="Right"/>
                        <TextBox x:Name="txtRestoreSrc" Margin="0,0,10,0"/>
                    </DockPanel>

                    <Button x:Name="btnRunRestore" Content="Run Restore" Height="36" Background="#FF27AE60" Foreground="White" FontWeight="Bold"/>
                </StackPanel>
            </TabItem>
        </TabControl>

        <GroupBox Header="Log" Grid.Row="1" Margin="0,10,0,0">
            <TextBox x:Name="txtLog" IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11"/>
        </GroupBox>
    </Grid>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$cmbUsersPrep      = $window.FindName("cmbUsersPrep")
$cmbUsersRestore   = $window.FindName("cmbUsersRestore")
$txtPrepDest       = $window.FindName("txtPrepDest")
$txtRestoreSrc     = $window.FindName("txtRestoreSrc")
$btnRefreshPrep    = $window.FindName("btnRefreshPrep")
$btnRefreshRestore = $window.FindName("btnRefreshRestore")
$btnBrowsePrepDest = $window.FindName("btnBrowsePrepDest")
$btnBrowseRestoreSrc = $window.FindName("btnBrowseRestoreSrc")
$btnRunPrep        = $window.FindName("btnRunPrep")
$btnRunRestore     = $window.FindName("btnRunRestore")
$script:LogBox     = $window.FindName("txtLog")

$txtPrepDest.Text = $DefaultBackupRoot

# ============================================================================
# EVENT WIRING
# ============================================================================
$btnRefreshPrep.Add_Click({ Update-ProfileComboBox -ComboBox $cmbUsersPrep; Write-Log "Profile list refreshed." })
$btnRefreshRestore.Add_Click({ Update-ProfileComboBox -ComboBox $cmbUsersRestore; Write-Log "Profile list refreshed." })

$btnBrowsePrepDest.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $txtPrepDest.Text
    if ($dlg.ShowDialog() -eq "OK") { $txtPrepDest.Text = $dlg.SelectedPath }
})

$btnBrowseRestoreSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $DefaultBackupRoot
    if ($dlg.ShowDialog() -eq "OK") { $txtRestoreSrc.Text = $dlg.SelectedPath }
})

$btnRunPrep.Add_Click({
    $profile = Get-SelectedProfile -ComboBox $cmbUsersPrep
    if (-not $profile) { [System.Windows.MessageBox]::Show("Select a profile first.", "No profile selected"); return }
    if (-not $profile.Exists) { [System.Windows.MessageBox]::Show("Profile folder doesn't exist - nothing to prep.", "Missing folder"); return }
    Invoke-ProfilePrep -Profile $profile -BackupRoot $txtPrepDest.Text
    Update-ProfileComboBox -ComboBox $cmbUsersPrep
})

$btnRunRestore.Add_Click({
    $profile = Get-SelectedProfile -ComboBox $cmbUsersRestore
    if (-not $profile) { [System.Windows.MessageBox]::Show("Select a profile first.", "No profile selected"); return }
    if ([string]::IsNullOrWhiteSpace($txtRestoreSrc.Text)) { [System.Windows.MessageBox]::Show("Choose a backup folder first.", "No backup folder"); return }
    Invoke-ProfileRestore -Profile $profile -BackupFolder $txtRestoreSrc.Text
})

# Initial population
Update-ProfileComboBox -ComboBox $cmbUsersPrep
Update-ProfileComboBox -ComboBox $cmbUsersRestore
Write-Log "Profile Rebuild Tool ready. Running as: $(whoami)"

# ============================================================================
# SHOW WINDOW
# ============================================================================
$window.ShowDialog() | Out-Null