# Read-only tools: status, folders, search and reading one message.

$DefaultLimit = 25
$MaxLimit = 100
$SearchSeconds = 45       # stop searching folders after this and return what was found
$MaxBodyChars = 30000

function Invoke-OutlookStatus {
    $outlook = Get-Outlook
    $session = $outlook.Session
    $stores = @()
    foreach ($store in $session.Stores) {
        $size = $null
        try { $size = [math]::Round((Get-Item -LiteralPath $store.FilePath).Length / 1GB, 2) } catch {}
        $stores += [ordered]@{ name = $store.DisplayName; file = $store.FilePath; size_gb = $size }
    }
    # Reading the current user's address is what triggers Outlook's
    # "A program is trying to access e-mail address information" prompt when no
    # antivirus is reported as active, so this doubles as the prompt test.
    return [ordered]@{
        outlook_version = $outlook.Version
        user            = $session.CurrentUser.Address
        stores          = $stores
    }
}

function Invoke-ListFolders($arguments) {
    $outlook = Get-Outlook
    $folders = Get-MailFolders $outlook (Get-Arg $arguments "store")
    $result = @()
    foreach ($entry in $folders) {
        $result += [ordered]@{
            store    = $entry.store.DisplayName
            folder   = $entry.path
            messages = $entry.folder.Items.Count
            unread   = $entry.folder.UnReadItemCount
        }
    }
    return , $result
}

function ConvertTo-DaslDate([string]$value, [int]$addDays) {
    # Filters compare dates in UTC; the user means local days.
    try {
        $day = [datetime]::ParseExact($value, "yyyy-MM-dd", $Invariant)
    } catch {
        throw "Dates must look like 2025-03-31, not '$value'."
    }
    return $day.AddDays($addDays).ToUniversalTime().ToString("yyyy-MM-dd HH:mm", $Invariant)
}

function New-SearchFilter($arguments, [bool]$indexed) {
    $parts = @()
    function Like([string]$property, [string]$text) {
        return "`"$property`" LIKE '%" + $text.Replace("'", "''") + "%'"
    }
    function Match([string]$property, [string]$text) {
        # Uses the Windows Search index: fast, matches whole words and word beginnings.
        return "`"$property`" ci_phrasematch '" + $text.Replace("'", "''") + "'"
    }
    $query = Get-Arg $arguments "query"
    if ($query) {
        # Scanning every body with LIKE takes minutes on a large PST, so use the index when Outlook has one.
        if ($indexed) {
            $parts += "(" + (Match "urn:schemas:httpmail:subject" $query) + " OR " +
                      (Match "urn:schemas:httpmail:textdescription" $query) + ")"
        } else {
            $parts += "(" + (Like "urn:schemas:httpmail:subject" $query) + " OR " +
                      (Like "urn:schemas:httpmail:textdescription" $query) + ")"
        }
    }
    $from = Get-Arg $arguments "from"
    if ($from) {
        $parts += "(" + (Like "urn:schemas:httpmail:fromname" $from) + " OR " +
                  (Like "http://schemas.microsoft.com/mapi/proptag/0x0C1F001F" $from) + ")"
    }
    $to = Get-Arg $arguments "to"
    if ($to) { $parts += Like "urn:schemas:httpmail:displayto" $to }
    $subject = Get-Arg $arguments "subject"
    if ($subject) { $parts += Like "urn:schemas:httpmail:subject" $subject }
    $since = Get-Arg $arguments "since"
    if ($since) { $parts += "`"urn:schemas:httpmail:datereceived`" >= '" + (ConvertTo-DaslDate $since 0) + "'" }
    $until = Get-Arg $arguments "until"
    if ($until) { $parts += "`"urn:schemas:httpmail:datereceived`" < '" + (ConvertTo-DaslDate $until 1) + "'" }
    if ((Get-Arg $arguments "unread") -eq $true) { $parts += "`"urn:schemas:httpmail:read`" = 0" }
    if ((Get-Arg $arguments "flagged") -eq $true) {
        # Follow-up flags set in current Outlook versions live in PR_TODO_ITEM_FLAGS; older ones in PR_FLAG_STATUS.
        $parts += "(`"http://schemas.microsoft.com/mapi/proptag/0x0E2B0003`" = 1 OR " +
                  "`"http://schemas.microsoft.com/mapi/proptag/0x10900003`" = 2)"
    }
    $category = Get-Arg $arguments "category"
    if ($category) { $parts += Like "urn:schemas-microsoft-com:office:office#Keywords" $category }
    if (-not $parts) { return $null }
    return "@SQL=" + ($parts -join " AND ")
}

function Invoke-Search($arguments) {
    $limit = Get-Arg $arguments "limit"
    if (-not $limit) { $limit = $DefaultLimit }
    $limit = [math]::Max(1, [math]::Min([int]$limit, $MaxLimit))
    $filters = @{}   # per store, since only some stores may be indexed
    $folderFilter = Get-Arg $arguments "folder"
    $onlyFlagged = (Get-Arg $arguments "flagged") -eq $true

    $outlook = Get-Outlook
    $folders = Get-MailFolders $outlook (Get-Arg $arguments "store")
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $found = New-Object System.Collections.ArrayList
    $searched = 0
    $incomplete = $false

    foreach ($entry in $folders) {
        if (-not (Test-Contains $entry.path $folderFilter)) { continue }
        if ($clock.Elapsed.TotalSeconds -gt $SearchSeconds) { $incomplete = $true; break }
        $searched++
        $storeId = $entry.store.StoreID
        if (-not $filters.ContainsKey($storeId)) {
            $indexed = $false
            try { $indexed = [bool]$entry.store.IsInstantSearchEnabled } catch {}
            $filters[$storeId] = New-SearchFilter $arguments $indexed
        }
        $filter = $filters[$storeId]
        $items = $entry.folder.Items
        if ($filter) {
            try { $items = $items.Restrict($filter) } catch { continue }
        }
        try { $items.Sort("[ReceivedTime]", $true) } catch {}
        # Each folder gives at most $limit of its newest matches; the best of all are kept below.
        $taken = 0
        $item = $items.GetFirst()
        while ($null -ne $item -and $taken -lt $limit) {
            $summary = Get-Summary $item $storeId $entry.store.DisplayName $entry.path
            # A completed flag keeps the to-do mark the filter matches on, so check it here.
            if (-not $onlyFlagged -or $summary.flag -eq "flagged") {
                [void]$found.Add($summary)
                $taken++
            }
            $item = $items.GetNext()
        }
    }

    $messages = @($found | Sort-Object { $_.received } -Descending | Select-Object -First $limit)
    $result = [ordered]@{
        found           = $messages.Count
        folders_checked = $searched
        messages        = $messages
    }
    if ($incomplete) {
        $result.note = "The search stopped after $SearchSeconds seconds before all folders were checked. " +
                       "Narrow it with since/until, folder or store to see everything."
    }
    return $result
}

function Get-VisibleAttachments($item) {
    # Leaves out pictures embedded in the body (logos, signatures), which Outlook itself hides.
    $list = @()
    $index = 0
    foreach ($attachment in $item.Attachments) {
        $index++
        if ((Get-Property $attachment $PropHidden) -eq $true) { continue }
        if ($attachment.Type -eq 6) { continue }   # olOLE: old embedded objects, cannot be saved as files
        $list += @{ index = $index; attachment = $attachment }
    }
    return , $list
}

function Invoke-GetMessage($arguments) {
    $id = Get-Arg $arguments "id"
    $storeId = Get-Arg $arguments "store_id"
    if (-not $id -or -not $storeId) { throw "Both id and store_id from outlook_search are required." }
    $outlook = Get-Outlook
    try {
        $item = $outlook.Session.GetItemFromID($id, $storeId)
    } catch {
        throw "No such message. It may have been moved or deleted in Outlook; search again."
    }
    $body = [string]$item.Body
    $truncated = $body.Length -gt $MaxBodyChars
    if ($truncated) { $body = $body.Substring(0, $MaxBodyChars) }
    $attachments = @()
    foreach ($entry in (Get-VisibleAttachments $item)) {
        $attachments += [ordered]@{
            index   = $entry.index
            name    = $entry.attachment.FileName
            size_kb = [math]::Round($entry.attachment.Size / 1KB)
        }
    }
    $result = [ordered]@{
        folder      = Get-FolderPath $item.Parent
        received    = Format-Time $item.ReceivedTime
        from        = Get-Sender $item
        to          = $item.To
        cc          = $item.CC
        subject     = $item.Subject
        categories  = $item.Categories
        attachments = $attachments
        body        = $body
    }
    if ($truncated) { $result.note = "The body was cut at $MaxBodyChars characters." }
    return $result
}
