# Drafts: write a message with files from this computer attached, and leave it open for the user.
# Nothing here sends: the draft is saved in Drafts and the user presses Send themselves.
# Keep this file ASCII-only: Windows PowerShell 5.1 reads scripts without a BOM as ANSI.

$MaxDraftFiles = 20
$MaxDraftBytes = 25MB     # all attachments together; most mail servers refuse larger messages
$MaxBodyChars = 100000

# Outlook data files and credential-like files that should never leave the computer by accident.
$BlockedAttachTypes = @(".pst", ".ost", ".nst", ".kdbx", ".pfx", ".p12", ".key", ".pem", ".ppk")

function Resolve-AttachFile([string]$path) {
    # Only an existing file on a drive letter (mapped network drives included), outside system folders.
    $hint = "Give the full path of a file, e.g. D:\Documents\Offer.pdf."
    if (-not $path) { throw $hint }
    $path = $path.Replace("/", "\")
    if ($path -notmatch '^[A-Za-z]:\\') { throw "'$path': $hint Paths like \\server\share are not allowed; use the mapped drive letter." }
    try { $full = [IO.Path]::GetFullPath($path) } catch { throw "'$path' is not a valid file path. $hint" }
    if ($full -notmatch '^[A-Za-z]:\\' -or $full.IndexOf(":", 2) -ge 0) { throw "'$path': device paths are not allowed." }
    foreach ($protected in Get-ProtectedFolders) {
        if ($full.StartsWith($protected, [StringComparison]::OrdinalIgnoreCase)) {
            throw "'$full' is in a system or application-data folder (Windows, Program Files, ProgramData, AppData). " +
                  "Only ordinary documents can be attached."
        }
    }
    $extension = [IO.Path]::GetExtension($full).ToLowerInvariant()
    if ($BlockedAttachTypes -contains $extension) {
        throw "'$full': $extension files (Outlook data files, keys and certificates) are never attached."
    }
    if (-not [IO.File]::Exists($full)) { throw "File not found: '$full'." }
    return Get-Item -LiteralPath $full
}

function Join-Recipients($value) {
    $list = @($value | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    foreach ($address in $list) {
        if ($address -match '[;,]') { throw "Give one recipient per list entry: '$address'." }
    }
    return ($list -join "; ")
}

function Invoke-CreateDraft($arguments) {
    $subject = [string](Get-Arg $arguments "subject")
    $body = [string](Get-Arg $arguments "body")
    if ($body.Length -gt $MaxBodyChars) { throw "The text is longer than $MaxBodyChars characters." }
    $to = Join-Recipients (Get-Arg $arguments "to")
    $cc = Join-Recipients (Get-Arg $arguments "cc")
    $bcc = Join-Recipients (Get-Arg $arguments "bcc")
    $display = (Get-Arg $arguments "open") -ne $false

    $files = @()
    $total = 0
    foreach ($path in @(Get-Arg $arguments "attachments" | Where-Object { $_ })) {
        $file = Resolve-AttachFile ([string]$path)
        if ($files | Where-Object { $_.FullName -eq $file.FullName }) { continue }
        $files += $file
        $total += $file.Length
    }
    if ($files.Count -gt $MaxDraftFiles) { throw "At most $MaxDraftFiles files can be attached to one message." }
    if ($total -gt $MaxDraftBytes) {
        throw ("The files are {0:N1} MB together; the limit is {1} MB. Send them in several messages or share a link." -f
            ($total / 1MB), ($MaxDraftBytes / 1MB))
    }

    $outlook = Get-Outlook
    $mail = $outlook.CreateItem(0)   # olMailItem
    $mail.Subject = $subject
    if ($to) { $mail.To = $to }
    if ($cc) { $mail.CC = $cc }
    if ($bcc) { $mail.BCC = $bcc }
    $mail.Body = $body
    foreach ($file in $files) { [void]$mail.Attachments.Add($file.FullName) }
    $mail.Save()
    if ($display) { $mail.Display($false) }

    return [ordered]@{
        saved_in    = "Drafts"
        opened      = $display
        sent        = $false
        to          = $to
        cc          = $cc
        bcc         = $bcc
        subject     = $subject
        attachments = @($files | ForEach-Object { [ordered]@{ file = $_.FullName; size_kb = [math]::Ceiling($_.Length / 1KB) } })
        note        = "The draft was not sent. The user checks it in Outlook and presses Send."
    }
}
