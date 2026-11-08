# Follow-up: which conversations wait for my answer, and which wait for someone else's.
#
# Outlook's own "replied" mark only knows replies sent from this Outlook, not those sent
# through the mail connector. So conversations are rebuilt from the messages themselves:
# Message-ID / In-Reply-To / References, and the subject plus the other party as a fallback.
# A conversation whose newest message is from me waits for others; one whose newest
# message is to me waits for me.

$FollowupSeconds = 50
$MaxFollowupItems = 6000
$SkipFolders = "Deleted Items|Junk|Drafts|Outbox|Spam|Trash|Infected|Lost & Found|Sync Issues|RSS"
$AutomaticSenders = "^((no-?reply|do-?not-?reply|bounces?)[\w.+-]*|mailer-daemon|postmaster|root|cron|daemon|alerts?|" +
                    "monitor\w*|[\w.-]*watch|nagios|zabbix|notifications?|updates?|newsletter|news|marketing)@"
# Headers that bulk-mail services, mailing lists and cron add. Many servers strip
# List-Unsubscribe, so the senders' own marks matter as much.
$AutomaticHeaders = "(?im)^(List-Unsubscribe|List-Id|Feedback-ID|X-Cron-Env|X-MarketoID|X-Mkt\w*|X-MSYS-API|X-KumoRef|" +
                    "X-Campaign\w*|X-Mailgun-\w+|X-SG-EID|X-MC-User|X-SES-Outgoing|X-CSA-Complaints):|" +
                    "(?im)^Precedence:\s*(bulk|list|junk)|(?im)^Auto-Submitted:\s*auto-"
# Mail a person wrote carries the name of the program they wrote it in; scripts rarely do.
$ClientHeaders = "(?im)^(X-Mailer|User-Agent):"
$ForwardSubject = "^\s*(fw|fwd)\s*:"

function Get-MyAddresses($outlook) {
    $mine = @{}
    try { $mine[$outlook.Session.CurrentUser.Address.ToLowerInvariant()] = $true } catch {}
    foreach ($account in $outlook.Session.Accounts) {
        try { if ($account.SmtpAddress) { $mine[$account.SmtpAddress.ToLowerInvariant()] = $true } } catch {}
    }
    foreach ($store in $outlook.Session.Stores) {
        if ($store.DisplayName -match "^[^@\s]+@[^@\s]+$") { $mine[$store.DisplayName.ToLowerInvariant()] = $true }
    }
    return $mine
}

function Get-Recipients($item) {
    $to = @()
    $cc = @()
    foreach ($recipient in $item.Recipients) {
        $address = $null
        try { $address = [string]$recipient.Address } catch {}
        if (-not $address) { continue }
        $address = $address.ToLowerInvariant()
        switch ([int]$recipient.Type) { 1 { $to += $address } 2 { $cc += $address } }
    }
    return @{ to = $to; cc = $cc }
}

function Get-MessageIds([string]$value) {
    if (-not $value) { return @() }
    return @([regex]::Matches($value, "<[^<>\s]+>") | ForEach-Object { $_.Value.ToLowerInvariant() })
}

function Test-Automatic($item, [string]$sender) {
    if ($sender -match $AutomaticSenders) { return $true }
    $headers = [string](Get-Property $item $PropHeaders)
    return $headers -and $headers -match $AutomaticHeaders
}

function Test-WrittenByMe($m) {
    # Mail from my address counts as my own only if it is in a Sent folder or came from a mail
    # program (e.g. a mail connector's BCC copies carry an X-Mailer header), not from a monitoring script.
    if ($m.in_sent) { return $true }
    $headers = [string](Get-Property $m.item $PropHeaders)
    return (-not $headers) -or ($headers -match $ClientHeaders -and $headers -notmatch $AutomaticHeaders)
}

function Get-SentPaths($outlook) {
    $paths = @{}
    foreach ($store in $outlook.Session.Stores) {
        try { $paths[(Get-FolderPath $store.GetDefaultFolder(5)).ToLowerInvariant()] = $true } catch {}   # olFolderSentMail
    }
    return $paths
}

function Get-Topic($item) {
    $topic = [string]$item.ConversationTopic
    if (-not $topic) { $topic = [string]$item.Subject }
    # Drop reply markers and tags such as "[SPAM]:" that servers add along the way.
    $topic = $topic -replace "^\s*((re|fw|fwd|aw|sv|vs|odp)\s*(\[\d+\])?\s*:\s*|\[[^\]]*\]\s*:?\s*)+", ""
    return $topic.Trim().ToLowerInvariant()
}

function Find-Root($parent, [string]$key) {
    while ($parent[$key] -ne $key) {
        $parent[$key] = $parent[$parent[$key]]
        $key = $parent[$key]
    }
    return $key
}

function Join-Keys($parent, [string]$a, [string]$b) {
    if (-not $parent.ContainsKey($a)) { $parent[$a] = $a }
    if (-not $parent.ContainsKey($b)) { $parent[$b] = $b }
    $ra = Find-Root $parent $a
    $rb = Find-Root $parent $b
    if ($ra -ne $rb) { $parent[$ra] = $rb }
}

function Get-FollowupMessages($outlook, [int]$days, [string]$storeFilter) {
    $mine = Get-MyAddresses $outlook
    $sentPaths = Get-SentPaths $outlook
    $since = (Get-Date).Date.AddDays(-$days).ToUniversalTime().ToString("yyyy-MM-dd HH:mm", $Invariant)
    $filter = "@SQL=`"urn:schemas:httpmail:datereceived`" >= '$since'"
    $clock = [Diagnostics.Stopwatch]::StartNew()
    $messages = New-Object System.Collections.ArrayList
    $incomplete = $false

    foreach ($entry in (Get-MailFolders $outlook $storeFilter)) {
        if ($entry.path -match $SkipFolders) { continue }
        if ($clock.Elapsed.TotalSeconds -gt $FollowupSeconds -or $messages.Count -ge $MaxFollowupItems) {
            $incomplete = $true
            break
        }
        $inSent = $sentPaths.ContainsKey($entry.path.ToLowerInvariant()) -or $entry.path -match "\\Sent( Items)?$"
        $inInbox = $entry.path -match "^[^\\]+\\Inbox$"
        try { $items = $entry.folder.Items.Restrict($filter) } catch { continue }
        $item = $items.GetFirst()
        while ($null -ne $item) {
            if ($item.MessageClass -like "IPM.Note*") {
                $sender = ([string]$item.SenderEmailAddress).ToLowerInvariant()
                $recipients = Get-Recipients $item
                $all = @($recipients.to) + @($recipients.cc)
                $others = @($all | Where-Object { -not $mine.ContainsKey($_) })
                $messageId = (Get-MessageIds ([string](Get-Property $item $PropMessageId)) | Select-Object -First 1)
                [void]$messages.Add(@{
                    item       = $item
                    store_id   = $entry.store.StoreID
                    store      = $entry.store.DisplayName
                    folder     = $entry.path
                    in_sent    = $inSent
                    in_inbox   = $inInbox
                    time       = [datetime]$item.ReceivedTime
                    sender     = $sender
                    from_me    = $mine.ContainsKey($sender)
                    to_me      = @($recipients.to | Where-Object { $mine.ContainsKey($_) }).Count -gt 0
                    cc_me      = @($recipients.cc | Where-Object { $mine.ContainsKey($_) }).Count -gt 0
                    others     = $others
                    message_id = $messageId
                    links      = @(Get-MessageIds ([string](Get-Property $item $PropInReplyTo))) +
                                 @(Get-MessageIds ([string](Get-Property $item $PropReferences)))
                    topic      = Get-Topic $item
                    replied    = @(102, 103) -contains (Get-Property $item $PropLastVerb)
                    automatic  = $null
                })
            }
            $item = $items.GetNext()
        }
    }
    return @{ messages = $messages; incomplete = $incomplete }
}

function Get-Conversations($messages) {
    $parent = @{}
    $index = 0
    foreach ($m in $messages) {
        $index++
        # The same message can be both in Sent Items and, as the connector's BCC copy, in the Inbox.
        $m.key = if ($m.message_id) { $m.message_id } else { "#$index" }
        Join-Keys $parent $m.key $m.key
        foreach ($link in $m.links) { Join-Keys $parent $m.key $link }
        if ($m.topic) {
            $parties = if ($m.from_me) { $m.others } else { @($m.sender) }
            foreach ($party in $parties) { Join-Keys $parent $m.key ("topic:" + $m.topic + "|" + $party) }
        }
    }
    $groups = @{}
    foreach ($m in $messages) {
        $root = Find-Root $parent $m.key
        if (-not $groups.ContainsKey($root)) { $groups[$root] = New-Object System.Collections.ArrayList }
        [void]$groups[$root].Add($m)
    }
    return $groups.Values
}

function Get-FollowupRow($m, [int]$count) {
    $item = $m.item
    return [ordered]@{
        id            = $item.EntryID
        store_id      = $m.store_id
        folder        = $m.folder
        received      = Format-Time $m.time
        waiting_days  = [int][math]::Floor(((Get-Date) - $m.time).TotalDays)
        from          = Get-Sender $item
        to            = $item.To
        subject       = $item.Subject
        conversation  = $count
        preview       = Get-Preview $item
    }
}

function Invoke-Followup($arguments, [bool]$waitingForMe) {
    $days = Get-Arg $arguments "days"
    if (-not $days) { $days = 30 }
    $days = [math]::Max(1, [math]::Min([int]$days, 365))
    $limit = Get-Arg $arguments "limit"
    if (-not $limit) { $limit = 50 }
    $limit = [math]::Max(1, [math]::Min([int]$limit, $MaxLimit))
    $includeCc = (Get-Arg $arguments "include_cc") -eq $true
    $allFolders = (Get-Arg $arguments "all_folders") -eq $true
    $includeForwards = (Get-Arg $arguments "include_forwards") -eq $true

    $outlook = Get-Outlook
    $scan = Get-FollowupMessages $outlook $days (Get-Arg $arguments "store")
    $rows = New-Object System.Collections.ArrayList
    $skippedAutomatic = 0

    foreach ($conversation in (Get-Conversations $scan.messages)) {
        # Only real correspondence counts: mail I wrote to someone else, or mail to me from someone
        # else that is in the Inbox (what rules sorted into other folders is usually not for answering).
        $real = @($conversation | Where-Object {
            if ($_.from_me) {
                $_.others.Count -gt 0 -and (Test-WrittenByMe $_)
            } else {
                ($_.to_me -or ($includeCc -and $_.cc_me)) -and ($_.in_inbox -or $allFolders)
            }
        })
        if (-not $real) { continue }
        $last = $real | Sort-Object { $_.time } | Select-Object -Last 1
        if ($waitingForMe) {
            if ($last.from_me -or $last.replied) { continue }
            if ($null -eq $last.automatic) { $last.automatic = Test-Automatic $last.item $last.sender }
            if ($last.automatic) { $skippedAutomatic++; continue }
        } else {
            if (-not $last.from_me) { continue }
            # A forward passes something on; it rarely asks for an answer.
            if (-not $includeForwards -and [string]$last.item.Subject -match $ForwardSubject) { continue }
        }
        [void]$rows.Add((Get-FollowupRow $last $conversation.Count))
    }

    $sorted = @($rows | Sort-Object { $_.received } -Descending | Select-Object -First $limit)
    $result = [ordered]@{ days = $days; found = $rows.Count; messages = $sorted }
    if ($waitingForMe -and $skippedAutomatic) {
        $result.skipped_automatic = $skippedAutomatic
    }
    if ($scan.incomplete) {
        $result.note = "Not all folders were checked in time. Use fewer days or the store argument for a full answer."
    }
    return $result
}
