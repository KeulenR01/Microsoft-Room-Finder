#Requires -Version 5.1
<#
.SYNOPSIS
    Automates Outlook Room Finder configuration for Exchange Online room mailboxes.

.DESCRIPTION
    Configuring rooms for Outlook Room Finder in Exchange Online is a repetitive and
    error-prone process. Each room mailbox requires Place metadata (City, Floor,
    Capacity, devices, accessibility, tags, etc.) to be set correctly before it
    appears and filters properly in Room Finder. On top of that, every room must be
    a member of at least one Room List (a distribution group of type RoomList) -
    without this, the room is invisible to Room Finder entirely.

    Doing this manually through the Exchange Admin Center or by running individual
    PowerShell commands per room is time-consuming, easy to get wrong, and offers
    no consistent overview of what has and has not been configured.

    This script solves that by guiding an administrator interactively through the
    full Room Finder setup in three structured steps:

      Step 1 - Inventory and select room mailboxes
               Retrieves all room mailboxes, displays them in a numbered list, and
               lets the operator select one or more rooms using flexible input
               (all, single numbers, comma-separated, ranges, or combinations).
               Each selected room is validated for type, address-list visibility,
               and SMTP uniqueness before proceeding.

      Step 2 - Verify and complete Place metadata
               Retrieves the current Place properties for each room via Get-Place
               and highlights missing fields. The operator is prompted to fill in
               missing values field by field. Fields with a suggested default can be
               accepted by typing / (slash). Tags are always offered as a numbered
               selection menu, even when all other fields are already filled.
               Changes are applied with Set-Place and verified immediately.

      Step 3 - Link or create Room Lists
               Retrieves all existing Room Lists and lets the operator assign each
               room to an existing list or create a new one. Uniqueness checks are
               performed before creating a new Room List. Membership is verified
               after adding via Get-DistributionGroupMember.

    At the end of each run a summary is printed per room. The operator can choose
    to run the script again for additional rooms without reconnecting.

.PARAMETER WhatIf
    Dry-run mode. Simulates all changes without executing Set-Place,
    New-DistributionGroup or Add-DistributionGroupMember.

.PARAMETER LogPath
    Path for the transcript log file. If omitted, a timestamped file is created
    in the script directory when the operator chooses to enable logging.

.PARAMETER ResultSize
    Number of room mailboxes to retrieve. Defaults to 'Unlimited'.

.REQUIREMENTS
    - ExchangeOnlineManagement module  (Install-Module ExchangeOnlineManagement)
    - Exchange Administrator role (required for Get/Set-Place, distribution groups)

.SOURCES
    Room Finder configuration : https://learn.microsoft.com/en-us/exchange/troubleshoot/outlook-issues/configure-room-finder-rooms-workspaces
    Get-Place / Set-Place     : https://learn.microsoft.com/en-us/powershell/module/exchange/get-place
    Distribution groups       : https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-distribution-groups/manage-distribution-groups
    Get-Recipient checks      : https://learn.microsoft.com/en-us/powershell/module/exchange/get-recipient
    EXO PowerShell module     : https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell

.EXAMPLE
    # Run interactively
    .\Set-RoomFinder.ps1

    # Dry-run - simulate all changes, write nothing
    .\Set-RoomFinder.ps1 -WhatIf

    # Write log to a custom path
    .\Set-RoomFinder.ps1 -LogPath "C:\Logs\roomfinder.log"

.NOTES
    Author  : <your name>
    Version : 1.0.0
#>

param(
    [switch]$WhatIf,
    [string]$LogPath    = '',
    [string]$ResultSize = 'Unlimited'
)

# ─── WhatIf mode: single central flag - no write actions anywhere in the script ──────
$DryRun = $WhatIf.IsPresent
if ($DryRun) { Write-Host "`n[DRYRUN] WhatIf active - no changes will be written.`n" -ForegroundColor Yellow }

$ErrorActionPreference = 'Continue'

# ─── Optional transcript log ──────────────────────────────────────────────────────────
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$logActive = $false

Write-Host "`nDo you want to create a log file?" -ForegroundColor Cyan
Write-Host "  Log will be saved to: $scriptDir" -ForegroundColor Gray
$logChoice = (Read-Host "  Create log file? (y/n)").Trim().ToLower()

if ($logChoice -eq 'y') {
    if ($LogPath -eq '') {
        $LogPath = Join-Path $scriptDir "RoomFinder_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
    }
    try {
        Start-Transcript -Path $LogPath -Append -Force | Out-Null
        Write-Host "[LOG] Transcript started: $LogPath`n" -ForegroundColor Cyan
        $logActive = $true
    } catch {
        Write-Warning "Could not start transcript: $_"
    }
}

# ─── Exchange Online connection ───────────────────────────────────────────────────────
# https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell
if (-not (Get-Command Get-Mailbox -ErrorAction SilentlyContinue)) {
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -ShowBanner:$false
}

# ─── Helper functions ─────────────────────────────────────────────────────────────────

function Write-OK   { param($m) Write-Host "  [ OK ] $m" -ForegroundColor Green  }
function Write-Warn { param($m) Write-Host "  [WARN] $m" -ForegroundColor Yellow }
function Write-Fail { param($m) Write-Host "  [FAIL] $m" -ForegroundColor Red    }

# Selection parser - accepts: all / 1 / 1,3 / 5-12 / combinations
function Read-Selection {
    param([int]$Max)
    do {
        $raw   = (Read-Host "Selection (all, 1, 1,3, 5-12 or combination)").Trim()
        $valid = $true
        $out   = [System.Collections.Generic.List[int]]::new()

        if ($raw -eq 'all') { return 0..($Max - 1) }

        foreach ($t in ($raw -split ',')) {
            $t = $t.Trim()
            if ($t -match '^\d+$') {
                $n = [int]$t
                if ($n -lt 1 -or $n -gt $Max) { $valid = $false; break }
                $out.Add($n - 1)
            } elseif ($t -match '^(\d+)-(\d+)$') {
                $a = [int]$Matches[1]; $b = [int]$Matches[2]
                if ($a -lt 1 -or $b -gt $Max -or $a -gt $b) { $valid = $false; break }
                $a..$b | ForEach-Object { $out.Add($_ - 1) }
            } else { $valid = $false; break }
        }
        if (-not $valid) { Write-Warn "Invalid input, please try again." }
    } while (-not $valid)

    return ($out | Sort-Object -Unique)
}

# Checks whether a name or SMTP address is already in use via Get-Recipient.
# OwnSmtp: the object's own address - finding itself is not a conflict.
# https://learn.microsoft.com/en-us/powershell/module/exchange/get-recipient
function Test-Unique {
    param([string]$Value, [string]$Type = 'smtp', [string]$OwnSmtp = '')
    $hit = Get-Recipient -Identity $Value -ErrorAction SilentlyContinue
    if ($hit -and $hit.PrimarySmtpAddress -ne $OwnSmtp) {
        Write-Warn "Conflict on identity: $($hit.PrimarySmtpAddress)"; return $false
    }
    if ($Type -eq 'smtp') {
        $hit = Get-Recipient -Filter "EmailAddresses -eq 'smtp:$Value'" -ErrorAction SilentlyContinue
        if ($hit -and $hit.PrimarySmtpAddress -ne $OwnSmtp) {
            Write-Warn "Conflict on EmailAddresses: $($hit.PrimarySmtpAddress)"; return $false
        }
    }
    return $true
}

# Returns a list of empty Place fields.
# Field order matches the display and prompt order in Step 2.
# City, Floor and Capacity are explicitly required by Microsoft for Room Finder.
# IsWheelChairAccessible: only missing when $null - false is a valid set value.
# Tags: missing when $null or empty array - always prompted separately.
# https://learn.microsoft.com/en-us/exchange/troubleshoot/outlook-issues/configure-room-finder-rooms-workspaces
function Get-MissingFields {
    param($Place)
    $fields = 'CountryOrRegion','City','Floor','Building','FloorLabel','Capacity',
              'AudioDeviceName','DisplayDeviceName','VideoDeviceName',
              'IsWheelChairAccessible','Tags'
    $fields | Where-Object {
        $v = $Place.$_
        $null -eq $v -or
        ($v -is [string] -and [string]::IsNullOrWhiteSpace($v)) -or
        ($v -is [array]  -and $v.Count -eq 0) -or
        ($_ -eq 'Capacity' -and $v -eq 0)
    }
}

# ====================================================================================
#  MAIN LOOP - operator can restart for additional rooms without reconnecting
# ====================================================================================
do {

# Summary table - reset on each iteration
$summary = [System.Collections.Generic.List[PSCustomObject]]::new()

# ====================================================================================
#  STEP 1 - Inventory and select room mailboxes
#  https://learn.microsoft.com/en-us/powershell/module/exchange/get-mailbox
# ====================================================================================
Write-Host "`n== STEP 1 - Room mailboxes ==" -ForegroundColor White

$rooms = @(Get-Mailbox -RecipientTypeDetails RoomMailbox -ResultSize $ResultSize)
if ($rooms.Count -eq 0) { Write-Fail "No room mailboxes found."; break }

# Numbered list
$rooms | ForEach-Object -Begin { $i = 1 } {
    $hidden = if ($_.HiddenFromAddressListsEnabled) { 'HIDDEN' } else { '' }
    Write-Host ("  {0,3}.  {1,-38}  {2,-40}  {3}" -f $i++, $_.DisplayName, $_.PrimarySmtpAddress, $hidden)
}

$selected = Read-Selection -Max $rooms.Count
$workList = [System.Collections.Generic.List[object]]::new()

foreach ($idx in $selected) {
    $r = $rooms[$idx]
    Write-Host "`n  > $($r.DisplayName)" -ForegroundColor Cyan

    $fail = $false
    if ($r.RecipientTypeDetails -ne 'RoomMailbox') { Write-Fail "Not a RoomMailbox"; $fail = $true } else { Write-OK "Type: RoomMailbox" }
    if ($r.HiddenFromAddressListsEnabled)           { Write-Warn "Hidden from address lists"; $fail = $true } else { Write-OK "Visible in address list" }

    # Pass OwnSmtp so the room does not flag itself as a conflict.
    # https://learn.microsoft.com/en-us/powershell/module/exchange/get-recipient
    $smtpOk = Test-Unique -Value $r.PrimarySmtpAddress -OwnSmtp $r.PrimarySmtpAddress
    if ($smtpOk) { Write-OK "SMTP unique" } else { Write-Fail "SMTP conflict"; $fail = $true }

    if ($fail) {
        $choice = Read-Host "  Issue found - [S]kip or [C]ontinue anyway?"
        if ($choice -notmatch '^[Cc]') { Write-Warn "Skipped."; continue }
        Write-Warn "Operator chose to continue despite issue(s)."
    }
    $workList.Add($r)
}

Write-Host "`n  $($workList.Count) room(s) proceeding to Step 2." -ForegroundColor Cyan

# ====================================================================================
#  STEP 2 - Verify and complete Place metadata
#  Field order: CountryOrRegion, City, Floor, Building, FloorLabel, Capacity,
#               AudioDeviceName, DisplayDeviceName, VideoDeviceName,
#               IsWheelChairAccessible, Tags
#  https://learn.microsoft.com/en-us/powershell/module/exchange/get-place
#  https://learn.microsoft.com/en-us/powershell/module/exchange/set-place
# ====================================================================================
Write-Host "`n== STEP 2 - Place metadata ==" -ForegroundColor White

# Fixed field order - identical to Get-MissingFields and the prompt loop below
$placeFields = 'CountryOrRegion','City','Floor','Building','FloorLabel','Capacity',
               'AudioDeviceName','DisplayDeviceName','VideoDeviceName',
               'IsWheelChairAccessible','Tags'

foreach ($r in $workList) {
    $smtp = $r.PrimarySmtpAddress
    Write-Host "`n  > $($r.DisplayName)" -ForegroundColor Cyan

    $entry = [PSCustomObject]@{
        DisplayName = $r.DisplayName
        Smtp        = $smtp
        MissingNa   = ''
        RoomList    = ''
        Added       = $false
        Error       = ''
    }

    try { $place = Get-Place -Identity $smtp -ErrorAction Stop }
    catch { Write-Fail "Get-Place failed: $_"; $entry.Error = "$_"; $summary.Add($entry); continue }

    # Tags is handled separately below - exclude it from the missing-fields loop
    $missing = @(Get-MissingFields -Place $place | Where-Object { $_ -ne 'Tags' })

    # Display current values - yellow = missing, gray = already set
    foreach ($f in $placeFields) {
        $v    = $place.$f
        $show = if ($null -eq $v -or ($v -is [string] -and $v -eq '') -or ($v -is [array] -and $v.Count -eq 0)) {
            '<empty>'
        } else { "$v" }
        $color = if ($f -in $missing) { 'Yellow' } else { 'Gray' }
        Write-Host ("    {0,-28} {1}" -f $f, $show) -ForegroundColor $color
    }

    # If all non-Tag fields are already filled: ask whether to edit anyway
    if ($missing.Count -eq 0) {
        Write-OK "All Place fields are filled in."
        $edit = (Read-Host "  Do you want to make changes anyway? (y/n)").Trim().ToLower()
        if ($edit -eq 'y') {
            $missing = $placeFields | Where-Object { $_ -ne 'Tags' }
        }
    }

    # ── Prompt for each missing field (Tags is always handled separately) ─────────
    $updates = @{}
    if ($missing.Count -gt 0) {
        Write-Warn "Missing fields: $($missing -join ', ')"
    }

    # Input convention for fields with a suggested value:
    #   Enter = skip (field remains unchanged)
    #   /     = accept the suggested default
    #   Text  = enter a custom value

    foreach ($f in $missing) {
        switch ($f) {

            'CountryOrRegion' {
                # Set-Place requires an ISO 3166 two-letter country code (e.g. NL, BE, DE)
                # https://learn.microsoft.com/en-us/powershell/module/exchange/set-place
                $v = (Read-Host "    CountryOrRegion  (ISO 3166 code e.g. NL/BE/DE, Enter=skip  /=NL)").Trim()
                if ($v -eq '')      { <# skip #> }
                elseif ($v -eq '/') { $updates[$f] = 'NL'; Write-Host "    -> NL" -ForegroundColor Gray }
                else                { $updates[$f] = $v.ToUpper() }
            }

            'Floor' {
                do {
                    $v = (Read-Host "    Floor  (integer, Enter=skip)").Trim()
                    if ($v -eq '') { break }
                } while ($v -notmatch '^\d+$' -and (Write-Warn "    Please enter a whole number.") -eq $null)
                if ($v -match '^\d+$') { $updates[$f] = [int]$v }
            }

            'Capacity' {
                do {
                    $v = (Read-Host "    Capacity  (integer, Enter=skip)").Trim()
                    if ($v -eq '') { break }
                } while ($v -notmatch '^\d+$' -and (Write-Warn "    Please enter a whole number.") -eq $null)
                if ($v -match '^\d+$') { $updates[$f] = [int]$v }
            }

            'AudioDeviceName' {
                # Suggested default: Speaker
                $v = (Read-Host "    AudioDeviceName  (Enter=skip  /=Speaker)").Trim()
                if ($v -eq '')      { <# skip #> }
                elseif ($v -eq '/') { $updates[$f] = 'Speaker';      Write-Host "    -> Speaker"      -ForegroundColor Gray }
                else                { $updates[$f] = $v }
            }

            'DisplayDeviceName' {
                # Suggested default: TV Screen
                $v = (Read-Host "    DisplayDeviceName  (Enter=skip  /=TV Screen)").Trim()
                if ($v -eq '')      { <# skip #> }
                elseif ($v -eq '/') { $updates[$f] = 'TV Screen';    Write-Host "    -> TV Screen"    -ForegroundColor Gray }
                else                { $updates[$f] = $v }
            }

            'VideoDeviceName' {
                # Suggested default: Teams Camera
                $v = (Read-Host "    VideoDeviceName  (Enter=skip  /=Teams Camera)").Trim()
                if ($v -eq '')      { <# skip #> }
                elseif ($v -eq '/') { $updates[$f] = 'Teams Camera'; Write-Host "    -> Teams Camera" -ForegroundColor Gray }
                else                { $updates[$f] = $v }
            }

            'IsWheelChairAccessible' {
                # Suggested default: false (not wheelchair accessible)
                # Type 'true' if the room IS wheelchair accessible
                $v = (Read-Host "    IsWheelChairAccessible  (Enter=skip  /=false  or type true)").Trim().ToLower()
                if ($v -eq '')          { <# skip #> }
                elseif ($v -eq '/')     { $updates[$f] = $false; Write-Host "    -> false" -ForegroundColor Gray }
                elseif ($v -eq 'true')  { $updates[$f] = $true;  Write-Host "    -> true"  -ForegroundColor Gray }
                elseif ($v -eq 'false') { $updates[$f] = $false; Write-Host "    -> false" -ForegroundColor Gray }
                else { Write-Warn "    Invalid input - skipped. Enter /, true or false." }
            }

            default {
                $v = (Read-Host "    $f  (Enter=skip)").Trim()
                if ($v) { $updates[$f] = $v }
            }
        }
    }

    # ── Tags - always prompted, even when all other fields are already filled ──────
    # Show current tags so the operator knows what is already set
    $currentTags = $place.Tags
    $currentTagsShow = if ($null -eq $currentTags -or ($currentTags -is [array] -and $currentTags.Count -eq 0)) {
        '<none>'
    } else { $currentTags -join ', ' }

    Write-Host "`n    Tags - would you like to add one or more tags?" -ForegroundColor Cyan
    Write-Host ("    Current tags : {0}" -f $currentTagsShow) -ForegroundColor Gray
    Write-Host "      1. Teams Room" -ForegroundColor Gray
    Write-Host "      2. Whiteboard" -ForegroundColor Gray
    Write-Host "      3. HDMI" -ForegroundColor Gray
    Write-Host "    Enter numbers separated by comma (e.g. 1,3) - Enter = skip" -ForegroundColor Gray

    $tagOptions = @('Teams Room', 'Whiteboard', 'HDMI')
    $tagInput   = (Read-Host "    Choice").Trim()

    if ($tagInput -ne '') {
        $chosenTags = $tagInput -split ',' | ForEach-Object {
            $n = $_.Trim() -as [int]
            if ($n -ge 1 -and $n -le $tagOptions.Count) { $tagOptions[$n - 1] }
            else { Write-Warn "    Invalid tag number: $($_.Trim()) - skipped." }
        } | Where-Object { $_ }

        if ($chosenTags.Count -gt 0) {
            $updates['Tags'] = $chosenTags
            Write-Host ("    -> Tags set to: {0}" -f ($chosenTags -join ', ')) -ForegroundColor Gray
        }
    } else {
        Write-Host "    -> Tags skipped." -ForegroundColor Gray
    }

    # Apply updates with Set-Place
    if ($updates.Count -gt 0) {
        if ($DryRun) {
            Write-Warn "[DRYRUN] Set-Place -Identity '$smtp' fields: $($updates.Keys -join ', ')"
        } else {
            try {
                # https://learn.microsoft.com/en-us/powershell/module/exchange/set-place
                Set-Place -Identity $smtp @updates -ErrorAction Stop
                Write-OK "Set-Place applied."

                # Verify by re-fetching and comparing
                $placeAfter = Get-Place -Identity $smtp -ErrorAction Stop
                foreach ($k in $updates.Keys) {
                    $after = $placeAfter.$k
                    if ("$after" -eq "$($updates[$k])") { Write-OK "$k : $after" }
                    else { Write-Warn "$k : expected '$($updates[$k])' but got '$after'" }
                }
                $entry.MissingNa = (Get-MissingFields -Place $placeAfter) -join ', '
            } catch { Write-Fail "Set-Place failed: $_"; $entry.Error += "$_" }
        }
    } else {
        Write-Warn "No values entered - Place unchanged."
    }

    $summary.Add($entry)
}

# ====================================================================================
#  STEP 3 - Link or create Room Lists
#  https://learn.microsoft.com/en-us/exchange/recipients-in-exchange-online/manage-distribution-groups/manage-distribution-groups
# ====================================================================================
Write-Host "`n== STEP 3 - Room Lists ==" -ForegroundColor White

$roomLists = @(Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' })

Write-Host "`n  Available Room Lists:"
$roomLists | ForEach-Object -Begin { $i = 1 } {
    Write-Host ("  {0,3}.  {1,-38}  {2}" -f $i++, $_.DisplayName, $_.PrimarySmtpAddress)
}
if ($roomLists.Count -eq 0) { Write-Warn "No Room Lists found." }

foreach ($r in $workList) {
    $smtp  = $r.PrimarySmtpAddress
    $entry = $summary | Where-Object { $_.Smtp -eq $smtp } | Select-Object -First 1
    if (-not $entry) { continue }

    Write-Host "`n  > $($r.DisplayName)" -ForegroundColor Cyan
    $choice = Read-Host "  [1] Add to existing Room List  [2] Create new  [3] Skip"

    $rlSmtp = $null

    if ($choice -eq '1') {
        if ($roomLists.Count -eq 0) { Write-Warn "No Room Lists available."; continue }
        do { $n = (Read-Host "  Number (1-$($roomLists.Count))") -as [int] }
        while ($null -eq $n -or $n -lt 1 -or $n -gt $roomLists.Count)
        $rlSmtp = $roomLists[$n - 1].PrimarySmtpAddress

    } elseif ($choice -eq '2') {
        $name     = (Read-Host "  Name for new Room List").Trim()
        $newSmtp  = (Read-Host "  PrimarySmtpAddress (Enter=auto)").Trim()

        # Uniqueness checks before creating
        # https://learn.microsoft.com/en-us/powershell/module/exchange/get-recipient
        if (-not (Test-Unique -Value $name -Type 'name')) { Write-Fail "Name already in use."; continue }
        if ($newSmtp -and -not (Test-Unique -Value $newSmtp)) { Write-Fail "SMTP already in use."; continue }

        if ($DryRun) {
            Write-Warn "[DRYRUN] New-DistributionGroup -Name '$name' -RoomList$(if ($newSmtp) { " -PrimarySmtpAddress '$newSmtp'" })"
            $rlSmtp = 'dryrun@example.com'
        } else {
            $params = @{ Name = $name; RoomList = $true }
            if ($newSmtp) { $params['PrimarySmtpAddress'] = $newSmtp }
            try {
                $created   = New-DistributionGroup @params -ErrorAction Stop
                $rlSmtp    = $created.PrimarySmtpAddress
                # Reload Room Lists so the new entry is available for subsequent rooms
                $roomLists = @(Get-DistributionGroup -ResultSize Unlimited | Where-Object { $_.RecipientTypeDetails -eq 'RoomList' })
                Write-OK "Room List created: $rlSmtp"
            } catch { Write-Fail "Creation failed: $_"; $entry.Error += "$_"; continue }
        }

    } else { Write-Warn "Skipped."; continue }

    $entry.RoomList = $rlSmtp

    if ($DryRun) {
        Write-Warn "[DRYRUN] Add-DistributionGroupMember -Identity '$rlSmtp' -Member '$smtp'"
        $entry.Added = $true
    } else {
        try {
            Add-DistributionGroupMember -Identity $rlSmtp -Member $smtp -ErrorAction Stop
            # Verify membership after adding
            $members = Get-DistributionGroupMember -Identity $rlSmtp -ResultSize Unlimited
            if ($members | Where-Object { $_.PrimarySmtpAddress -eq $smtp }) {
                Write-OK "Membership verified."
                $entry.Added = $true
            } else { Write-Fail "Membership verification failed." }
        } catch {
            if ($_ -match 'already a member') {
                Write-OK "Already a member of '$rlSmtp'."
                $entry.Added = $true
                $change = (Read-Host "  Do you want to change the Room List? (y/n)").Trim().ToLower()
                if ($change -eq 'y') {
                    Write-Warn "Remove the room manually from '$rlSmtp' and re-run the script, or choose a different list."
                }
            } else { Write-Fail "Add failed: $_"; $entry.Error += "$_" }
        }
    }
}

# ====================================================================================
#  STEP 4 - Summary
# ====================================================================================
Write-Host "`n== SUMMARY ==" -ForegroundColor White

foreach ($e in $summary) {
    Write-Host ""
    Write-Host ("Resource: {0}" -f $e.DisplayName) -ForegroundColor White
    if ($e.MissingNa) {
        Write-Host ("Missing values : {0}" -f $e.MissingNa) -ForegroundColor Red
    } else {
        Write-Host "Missing values : none" -ForegroundColor Green
        Write-Host "All required properties have been set on the resource account." -ForegroundColor Green
    }
}

if ($DryRun) { Write-Host "`n[DRYRUN] No changes were written." -ForegroundColor Yellow }

# ─── Repeat? ──────────────────────────────────────────────────────────────────────
Write-Host ""
$again = (Read-Host "Do you want to run through this script again for other room(s)? (y/n)").Trim().ToLower()

} while ($again -eq 'y')

# ─── Exit ─────────────────────────────────────────────────────────────────────────
Write-Host "`nScript finished." -ForegroundColor Cyan
if ($logActive) {
    Write-Host "[LOG] Saved to: $LogPath`n" -ForegroundColor Cyan
    Stop-Transcript | Out-Null
}
