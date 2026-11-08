# Tools that tidy local mail: create folders, move messages, mark them.
# Nothing here deletes or sends anything, and every change can be undone in Outlook.

# Default folders that messages must never be moved into (or below):
# 3 olFolderDeletedItems, 23 olFolderJunk, 4 olFolderOutbox (would send), 20 olFolderSyncIssues.
$ProtectedDefaultFolders = @(3, 23, 4, 20)

function Test-SameEntryId($outlook, [string]$a, [string]$b) {
    if (-not $a -or -not $b) { return $false }
    if ([string]::Equals($a, $b, [StringComparison]::OrdinalIgnoreCase)) { return $true }
    try { return [bool]$outlook.Session.CompareEntryIDs($a, $b) } catch { return $false }
}

function Assert-AllowedTarget($outlook, $folder) {
    # Refuses Deleted Items, Junk, Outbox and Sync Issues of the folder's own data file, and folders under them.
    $store = $null
    try { $store = $folder.Store } catch {}
    $protected = @()
    if ($store) {
        foreach ($kind in $ProtectedDefaultFolders) {
            try { $protected += [string]$store.GetDefaultFolder($kind).EntryID } catch {}
        }
    }
    if (-not $protected) { return }
    $current = $folder
    for ($depth = 0; $current -and $depth -lt 64; $depth++) {
        $entryId = $null
        try { $entryId = [string]$current.EntryID } catch { break }
        foreach ($id in $protected) {
            if (Test-SameEntryId $outlook $entryId $id) {
                throw "'$(Get-FolderPath $folder)' is in Deleted Items, Junk E-mail, Outbox or Sync Issues. " +
                      "The plugin never deletes mail (and never sends it), so it does not move messages there. " +
                      "To clean up, move them to an ordinary folder that the user empties themselves."
            }
        }
        # Climb to the parent folder; the root folder's parent is the session, not a folder (Class 2 = olFolder).
        try {
            $parent = $current.Parent
            if ($null -eq $parent -or $parent.Class -ne 2) { break }
            $current = $parent
        } catch { break }
    }
}

function Invoke-CreateFolder($arguments) {
    $name = Get-Arg $arguments "name"
    if (-not $name) { throw "A folder name is required." }
    if ($name -match '[\\/]') { throw "The name cannot contain \ or /. Create nested folders one level at a time." }
    $outlook = Get-Outlook
    $parent = Find-Folder $outlook (Get-Arg $arguments "parent")
    Assert-AllowedTarget $outlook $parent
    foreach ($sub in $parent.Folders) {
        if ([string]::Equals($sub.Name, $name, [StringComparison]::OrdinalIgnoreCase)) {
            return [ordered]@{ folder = Get-FolderPath $sub; created = $false; note = "The folder already exists." }
        }
    }
    $folder = $parent.Folders.Add($name)
    return [ordered]@{ folder = Get-FolderPath $folder; created = $true }
}

function Invoke-MoveMessages($arguments) {
    $outlook = Get-Outlook
    $target = Find-Folder $outlook (Get-Arg $arguments "folder")
    if ($target.DefaultItemType -ne 0) { throw "'$(Get-FolderPath $target)' is not a mail folder." }
    Assert-AllowedTarget $outlook $target
    $items = Get-Messages $outlook (Get-Arg $arguments "messages")
    $targetStore = $target.Store
    $moved = @()
    $failed = @()
    foreach ($item in $items) {
        $subject = $item.Subject
        try {
            # Move returns the item in its new place, with a new id.
            $new = $item.Move($target)
            $moved += [ordered]@{ id = $new.EntryID; store_id = $targetStore.StoreID; subject = $subject }
        } catch {
            $failed += [ordered]@{ subject = $subject; error = $_.Exception.Message }
        }
    }
    $result = [ordered]@{ folder = Get-FolderPath $target; moved = $moved.Count; messages = $moved }
    if ($failed) { $result.failed = $failed }
    return $result
}

function Split-Categories([string]$value) {
    # Outlook joins categories with the Windows list separator (", " or "; ").
    if (-not $value) { return @() }
    return @($value -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

function Invoke-MarkMessages($arguments) {
    $read = Get-Arg $arguments "read"
    $flag = Get-Arg $arguments "flag"
    $add = @(Get-Arg $arguments "add_categories" | Where-Object { $_ })
    $remove = @(Get-Arg $arguments "remove_categories" | Where-Object { $_ })
    if ($null -eq $read -and -not $flag -and -not $add -and -not $remove) {
        throw "Say what to change: read, flag, add_categories or remove_categories."
    }
    if ($flag -and $flag -notin @("flag", "complete", "clear")) { throw "flag must be 'flag', 'complete' or 'clear'." }
    $separator = [Globalization.CultureInfo]::CurrentCulture.TextInfo.ListSeparator + " "

    $outlook = Get-Outlook
    $items = Get-Messages $outlook (Get-Arg $arguments "messages")
    $changed = 0
    $failed = @()
    foreach ($item in $items) {
        try {
            if ($null -ne $read) { $item.UnRead = -not [bool]$read }
            switch ($flag) {
                "flag" { $item.MarkAsTask(4) }   # olMarkNoDate
                "complete" { $item.FlagStatus = 1 }   # olFlagComplete
                "clear" { $item.ClearTaskFlag() }
            }
            if ($add -or $remove) {
                $categories = New-Object System.Collections.ArrayList
                foreach ($category in (Split-Categories $item.Categories)) {
                    if ($remove -notcontains $category) { [void]$categories.Add($category) }
                }
                foreach ($category in $add) {
                    if ($categories -notcontains $category) { [void]$categories.Add($category) }
                }
                $item.Categories = $categories -join $separator
            }
            $item.Save()
            $changed++
        } catch {
            $failed += [ordered]@{ subject = $item.Subject; error = $_.Exception.Message }
        }
    }
    $result = [ordered]@{ changed = $changed }
    if ($failed) { $result.failed = $failed }
    return $result
}
