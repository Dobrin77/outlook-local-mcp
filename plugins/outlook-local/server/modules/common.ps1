# Shared helpers: Outlook access, arguments, folders and message summaries.
# Keep this file ASCII-only: Windows PowerShell 5.1 reads scripts without a BOM as ANSI.

$Invariant = [Globalization.CultureInfo]::InvariantCulture
$PreviewChars = 200
$MaxBatch = 200          # most messages one change may touch

# MAPI properties read through PropertyAccessor.
$PropHeaders = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"     # PR_TRANSPORT_MESSAGE_HEADERS
$PropMessageId = "http://schemas.microsoft.com/mapi/proptag/0x1035001F"   # PR_INTERNET_MESSAGE_ID
$PropInReplyTo = "http://schemas.microsoft.com/mapi/proptag/0x1042001F"   # PR_IN_REPLY_TO_ID
$PropReferences = "http://schemas.microsoft.com/mapi/proptag/0x1039001F"  # PR_INTERNET_REFERENCES
$PropLastVerb = "http://schemas.microsoft.com/mapi/proptag/0x10810003"    # PR_LAST_VERB_EXECUTED
$PropHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"      # PR_ATTACHMENT_HIDDEN

function Get-Outlook {
    # Attaches to the running Outlook, or starts it in the background.
    try {
        return New-Object -ComObject Outlook.Application
    } catch {
        throw "Outlook is not available on this computer. Classic Outlook (not 'New Outlook') is required."
    }
}

function Test-Contains([string]$text, [string]$part) {
    return (-not $part) -or ($text -and $text.IndexOf($part, [StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Get-Arg($arguments, [string]$name) {
    if ($null -eq $arguments) { return $null }
    $value = $arguments.$name
    if ($value -is [string]) { $value = $value.Trim() }
    return $value
}

function Get-Property($item, [string]$property) {
    try { return $item.PropertyAccessor.GetProperty($property) } catch { return $null }
}

function Get-MailFolders($outlook, [string]$storeFilter) {
    $list = New-Object System.Collections.ArrayList
    foreach ($store in $outlook.Session.Stores) {
        if (-not (Test-Contains $store.DisplayName $storeFilter)) { continue }
        try { $root = $store.GetRootFolder() } catch { continue }
        Add-MailFolders $root $store $list
    }
    return , $list
}

function Add-MailFolders($folder, $store, $list) {
    foreach ($sub in $folder.Folders) {
        if ($sub.DefaultItemType -eq 0) {   # olMailItem
            [void]$list.Add(@{ folder = $sub; store = $store; path = $sub.FolderPath.TrimStart("\") })
        }
        Add-MailFolders $sub $store $list
    }
}

function Find-Folder($outlook, [string]$path) {
    # A folder is named by its full path as outlook_list_folders shows it, e.g. "Mailbox\Inbox\Invoices".
    if (-not $path) { throw "A folder path is required, e.g. 'Mailbox\Inbox'." }
    $wanted = $path.Trim().Trim("\").Replace("/", "\")
    foreach ($store in $outlook.Session.Stores) {
        try { $root = $store.GetRootFolder() } catch { continue }
        $found = Find-FolderUnder $root $wanted
        if ($found) { return $found }
    }
    throw "No folder '$path'. Use the full path from outlook_list_folders, e.g. 'Mailbox\Inbox'."
}

function Find-FolderUnder($folder, [string]$wanted) {
    foreach ($sub in $folder.Folders) {
        $path = $sub.FolderPath.TrimStart("\")
        if ([string]::Equals($path, $wanted, [StringComparison]::OrdinalIgnoreCase)) { return $sub }
        if ($wanted.StartsWith($path + "\", [StringComparison]::OrdinalIgnoreCase)) {
            $found = Find-FolderUnder $sub $wanted
            if ($found) { return $found }
        }
    }
    return $null
}

function Get-FolderPath($folder) {
    try { return $folder.FolderPath.TrimStart("\") } catch { return $null }
}

function Get-Messages($outlook, $messages) {
    # Resolves [{id, store_id}, ...] from outlook_search into Outlook items.
    $list = @($messages)
    if (-not $messages -or $list.Count -eq 0) { throw "Give the messages as a list of {id, store_id} from outlook_search." }
    if ($list.Count -gt $MaxBatch) { throw "At most $MaxBatch messages at a time; split the change into smaller parts." }
    $items = New-Object System.Collections.ArrayList
    foreach ($message in $list) {
        $id = Get-Arg $message "id"
        $storeId = Get-Arg $message "store_id"
        if (-not $id -or -not $storeId) { throw "Every message needs both id and store_id from outlook_search." }
        try {
            [void]$items.Add($outlook.Session.GetItemFromID($id, $storeId))
        } catch {
            throw "Message $id was not found. It may have been moved or deleted in Outlook; search again."
        }
    }
    return , $items
}

function Get-Preview($item) {
    try { $body = [string]$item.Body } catch { return "" }
    $body = ($body -replace "\s+", " ").Trim()
    if ($body.Length -gt $PreviewChars) { $body = $body.Substring(0, $PreviewChars) + "..." }
    return $body
}

function Get-Sender($item) {
    $name = $item.SenderName
    $address = $item.SenderEmailAddress
    if ($address -and $address -ne $name) { return "$name <$address>" }
    return $name
}

function Format-Time($value) {
    try { return ([datetime]$value).ToString("yyyy-MM-dd HH:mm", $Invariant) } catch { return $null }
}

function Get-Summary($item, [string]$storeId, [string]$storeName, [string]$folderPath) {
    $flag = "none"
    try {
        switch ([int]$item.FlagStatus) { 1 { $flag = "complete" } 2 { $flag = "flagged" } }
    } catch {}
    return [ordered]@{
        id          = $item.EntryID
        store_id    = $storeId
        store       = $storeName
        folder      = $folderPath
        received    = Format-Time $item.ReceivedTime
        from        = Get-Sender $item
        to          = $item.To
        subject     = $item.Subject
        unread      = [bool]$item.UnRead
        flag        = $flag
        categories  = $item.Categories
        attachments = $item.Attachments.Count
        preview     = Get-Preview $item
    }
}

function Get-StoreOf($item) {
    try { return $item.Parent.Store } catch { return $null }
}
