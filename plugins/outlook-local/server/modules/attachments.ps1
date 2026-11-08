# Attachments: save them to a folder on disk, or hand one to Claude to read.

$MaxSaveFiles = 500
$MaxReadBytes = 10MB
$MaxTextChars = 60000

$ImageTypes = @{ ".png" = "image/png"; ".jpg" = "image/jpeg"; ".jpeg" = "image/jpeg"; ".gif" = "image/gif"; ".webp" = "image/webp" }
$TextTypes = @(".txt", ".csv", ".log", ".xml", ".json", ".md", ".html", ".htm", ".ics", ".vcf")
$MaxZipEntryBytes = 20MB

# Files that run code when opened: saved only when the caller asks for that extension by name.
$RunnableTypes = @(".exe", ".com", ".bat", ".cmd", ".vbs", ".vbe", ".js", ".jse", ".wsf", ".wsh", ".hta", ".lnk",
    ".ps1", ".psm1", ".scr", ".msi", ".msp", ".dll", ".cpl", ".jar", ".reg", ".url", ".pif")
$ReservedNamePattern = '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])\s*(\..*)?$'

function Get-SafeFileName([string]$name) {
    foreach ($char in [IO.Path]::GetInvalidFileNameChars()) { $name = $name.Replace([string]$char, "_") }
    $name = $name.Trim().TrimEnd(".").Trim()
    if (-not $name) { $name = "attachment" }
    # Windows device names (CON, NUL, COM1, ...) cannot be ordinary files, with or without an extension.
    if ($name -match $ReservedNamePattern) { $name = "_" + $name }
    return $name
}

function Get-ProtectedFolders {
    # System and autostart folders that attachments must never be written into.
    $paths = @(
        $env:SystemRoot, $env:windir, $env:ProgramFiles, ${env:ProgramFiles(x86)}, $env:ProgramW6432,
        $env:ProgramData, $env:APPDATA, $env:LOCALAPPDATA,
        [Environment]::GetFolderPath("Windows"),
        [Environment]::GetFolderPath("ProgramFiles"),
        [Environment]::GetFolderPath("ProgramFilesX86"),
        [Environment]::GetFolderPath("CommonApplicationData"),
        [Environment]::GetFolderPath("ApplicationData"),
        [Environment]::GetFolderPath("LocalApplicationData"),
        [Environment]::GetFolderPath("Startup"),
        [Environment]::GetFolderPath("CommonStartup")
    )
    $list = @()
    foreach ($path in $paths) {
        if (-not $path) { continue }
        try { $list += [IO.Path]::GetFullPath($path).TrimEnd("\") + "\" } catch {}
    }
    return $list
}

function Resolve-SaveDirectory([string]$directory) {
    # Only an ordinary folder on a local drive letter: no network shares, device paths or system folders.
    $hint = "Give a full folder path on a local drive, e.g. D:\Accounting\Invoices."
    if (-not $directory) { throw $hint }
    $directory = $directory.Replace("/", "\")
    # Rejects \\server\share, \\?\..., \\.\..., relative paths and drive-relative paths like 'C:folder'.
    if ($directory -notmatch '^[A-Za-z]:\\') { throw "$hint Network and device paths are not allowed." }
    try { $full = [IO.Path]::GetFullPath($directory) } catch { throw "'$directory' is not a valid folder path. $hint" }
    if ($full -notmatch '^[A-Za-z]:\\' -or $full.IndexOf(":", 2) -ge 0) {
        throw "$hint Network and device paths are not allowed."
    }
    foreach ($segment in $full.Substring(3).Split("\")) {
        if ($segment -and $segment -match $ReservedNamePattern) { throw "'$full' uses a Windows device name ($segment). $hint" }
    }
    $withSeparator = $full.TrimEnd("\") + "\"
    foreach ($protected in Get-ProtectedFolders) {
        if ($withSeparator.StartsWith($protected, [StringComparison]::OrdinalIgnoreCase)) {
            throw "Attachments cannot be saved in '$full': Windows, Program Files, ProgramData, AppData and " +
                  "Startup folders are protected. Choose an ordinary folder, e.g. D:\Accounting\Invoices."
        }
    }
    if ([IO.File]::Exists($full)) { throw "'$full' is a file, not a folder." }
    if ($full.Length -gt 3) { $full = $full.TrimEnd("\") }
    return $full
}

function Get-FreePath([string]$directory, [string]$name) {
    # Never overwrites: "file.pdf" becomes "file (2).pdf" if the name is taken.
    $path = [IO.Path]::Combine($directory, $name)
    $stem = [IO.Path]::GetFileNameWithoutExtension($name)
    $extension = [IO.Path]::GetExtension($name)
    $n = 2
    while (Test-Path -LiteralPath $path) {
        $path = [IO.Path]::Combine($directory, "$stem ($n)$extension")
        $n++
    }
    return $path
}

function Save-Attachment($attachment, [string]$path) {
    if ($attachment.Type -eq 5) {
        # An attached e-mail: save it as a .msg that Outlook opens.
        if (-not $path.EndsWith(".msg", [StringComparison]::OrdinalIgnoreCase)) { $path += ".msg" }
    }
    $attachment.SaveAsFile($path)
    return $path
}

function Invoke-SaveAttachments($arguments) {
    $directory = Resolve-SaveDirectory ([string](Get-Arg $arguments "directory"))
    $only = @(Get-Arg $arguments "extensions" | Where-Object { $_ } | ForEach-Object { "." + ([string]$_).Trim().TrimStart(".").ToLowerInvariant() })
    $datePrefix = (Get-Arg $arguments "date_prefix") -ne $false

    $outlook = Get-Outlook
    $items = Get-Messages $outlook (Get-Arg $arguments "messages")
    if (-not (Test-Path -LiteralPath $directory -PathType Container)) { [void][IO.Directory]::CreateDirectory($directory) }

    $saved = @()
    $skipped = @()
    $failed = @()
    foreach ($item in $items) {
        $prefix = ""
        if ($datePrefix) {
            try { $prefix = ([datetime]$item.ReceivedTime).ToString("yyyy-MM-dd", $Invariant) + " " } catch {}
        }
        foreach ($entry in (Get-VisibleAttachments $item)) {
            $attachment = $entry.attachment
            $name = Get-SafeFileName $attachment.FileName
            # An attached e-mail is saved as .msg; name it so before checking types and free names.
            if ($attachment.Type -eq 5 -and -not $name.EndsWith(".msg", [StringComparison]::OrdinalIgnoreCase)) { $name += ".msg" }
            $extension = [IO.Path]::GetExtension($name).ToLowerInvariant()
            if ($only -and $only -notcontains $extension) { continue }
            if ($RunnableTypes -contains $extension -and $only -notcontains $extension) {
                $skipped += [ordered]@{ name = $name; subject = $item.Subject
                    reason = "$extension files can run code; list '$($extension.TrimStart('.'))' in extensions to save them anyway." }
                continue
            }
            if ($saved.Count -ge $MaxSaveFiles) { throw "Stopped after $MaxSaveFiles files; save the rest in a second step." }
            try {
                $target = Get-FreePath $directory ($prefix + $name)
                if (-not [string]::Equals([IO.Path]::GetDirectoryName($target).TrimEnd("\"), $directory.TrimEnd("\"), [StringComparison]::OrdinalIgnoreCase)) {
                    throw "Refused: the file name would leave the chosen folder."
                }
                $path = Save-Attachment $attachment $target
                $saved += [ordered]@{ file = $path; from_subject = $item.Subject }
            } catch {
                $failed += [ordered]@{ name = $name; subject = $item.Subject; error = $_.Exception.Message }
            }
        }
    }
    $result = [ordered]@{ directory = $directory; saved = $saved.Count; files = $saved }
    if ($skipped) { $result.skipped = $skipped }
    if ($failed) { $result.failed = $failed }
    return $result
}

function ConvertFrom-Markup([string]$xml, [string]$paragraphTag) {
    # Crude but dependable text from Office XML or HTML: paragraph ends become new lines, tags go.
    $text = $xml -replace "</$paragraphTag>", "`n"
    $text = $text -replace "<(br|w:br|w:tab)[^>]*/?>", " "
    $text = $text -replace "<[^>]+>", ""
    $text = [Net.WebUtility]::HtmlDecode($text)
    return ($text -replace "[ \t]+", " " -replace "(\r?\n\s*){3,}", "`n`n").Trim()
}

function Read-ZipEntry($zip, [string]$name) {
    # Office files are zip archives; a tiny file can unpack to gigabytes, so every part is capped.
    $entry = $zip.GetEntry($name)
    if (-not $entry) { return $null }
    if ($entry.Length -gt $MaxZipEntryBytes) { throw "The attachment is too large to read here (part '$name' unpacks to more than 20 MB)." }
    $reader = New-Object IO.StreamReader($entry.Open(), [Text.Encoding]::UTF8)
    try {
        # Do not trust the size in the zip header alone: stop reading once the cap is passed.
        $out = New-Object Text.StringBuilder
        $buffer = New-Object char[] 65536
        while (($count = $reader.Read($buffer, 0, $buffer.Length)) -gt 0) {
            [void]$out.Append($buffer, 0, $count)
            if ($out.Length -gt $MaxZipEntryBytes) { throw "The attachment is too large to read here (part '$name' unpacks to more than 20 MB)." }
        }
        return $out.ToString()
    } finally { $reader.Dispose() }
}

function ConvertTo-SafeXml([string]$text) {
    # Parse XML without DTDs or external entities (no XXE, no entity expansion bombs).
    if (-not $text) { return $null }
    $settings = New-Object Xml.XmlReaderSettings
    $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
    $settings.XmlResolver = $null
    $reader = [Xml.XmlReader]::Create((New-Object IO.StringReader($text)), $settings)
    try {
        $document = New-Object Xml.XmlDocument
        $document.XmlResolver = $null
        $document.Load($reader)
    } finally { $reader.Dispose() }
    return , $document
}

function Read-Docx([string]$path) {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try { return ConvertFrom-Markup (Read-ZipEntry $zip "word/document.xml") "w:p" } finally { $zip.Dispose() }
}

function Read-Xlsx([string]$path) {
    # Every sheet as rows of tab-separated cells; formulas show their last calculated value.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($path)
    try {
        $shared = @()
        $sharedXml = Read-ZipEntry $zip "xl/sharedStrings.xml"
        if ($sharedXml) {
            foreach ($si in (ConvertTo-SafeXml $sharedXml).sst.si) {
                $shared += (@($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText }) -join "")
            }
        }
        $names = @{}
        $workbook = ConvertTo-SafeXml (Read-ZipEntry $zip "xl/workbook.xml")
        $index = 0
        if ($workbook) {
            foreach ($sheet in $workbook.workbook.sheets.sheet) { $index++; $names["xl/worksheets/sheet$index.xml"] = $sheet.name }
        }
        $out = New-Object Text.StringBuilder
        foreach ($entry in ($zip.Entries | Where-Object { $_.FullName -like "xl/worksheets/sheet*.xml" } | Sort-Object FullName)) {
            # The answer is cut at $MaxTextChars anyway; do not unpack sheets nobody will see.
            if ($out.Length -gt $MaxTextChars) { break }
            $title = $names[$entry.FullName]
            if (-not $title) { $title = $entry.Name }
            [void]$out.AppendLine("## $title")
            $sheetXml = ConvertTo-SafeXml (Read-ZipEntry $zip $entry.FullName)
            if (-not $sheetXml) { continue }
            foreach ($row in $sheetXml.worksheet.sheetData.row) {
                $cells = @()
                foreach ($cell in $row.c) {
                    $value = $null
                    if ($cell.t -eq "s") { $value = $shared[[int]$cell.v] }
                    elseif ($cell.t -eq "inlineStr") { $value = $cell.is.InnerText }
                    else { $value = $cell.v }
                    $cells += [string]$value
                }
                if (($cells -join "").Trim()) { [void]$out.AppendLine($cells -join "`t") }
            }
            [void]$out.AppendLine()
        }
        return $out.ToString().Trim()
    } finally { $zip.Dispose() }
}

function Invoke-ReadAttachment($arguments) {
    $id = Get-Arg $arguments "id"
    $storeId = Get-Arg $arguments "store_id"
    $index = [int](Get-Arg $arguments "index")
    if (-not $id -or -not $storeId -or $index -lt 1) {
        throw "id, store_id and the attachment index from outlook_get_message are required."
    }
    $outlook = Get-Outlook
    try { $item = $outlook.Session.GetItemFromID($id, $storeId) } catch { throw "No such message; search again." }
    if ($index -gt $item.Attachments.Count) { throw "The message has no attachment number $index." }
    $attachment = $item.Attachments.Item($index)
    $name = Get-SafeFileName $attachment.FileName
    if ($attachment.Size -gt $MaxReadBytes) {
        throw "$name is larger than 10 MB. Save it with outlook_save_attachments and open it yourself."
    }
    $extension = [IO.Path]::GetExtension($name).ToLowerInvariant()

    $directory = Join-Path ([IO.Path]::GetTempPath()) ("outlook-local-" + [guid]::NewGuid().ToString("N"))
    [void](New-Item -ItemType Directory -Path $directory)
    try {
        $path = Save-Attachment $attachment (Join-Path $directory $name)
        $header = "Attachment '$name' from the message '$($item.Subject)':"
        if ($ImageTypes.ContainsKey($extension)) {
            $data = [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))
            return @{ __content = @(
                @{ type = "text"; text = $header },
                @{ type = "image"; data = $data; mimeType = $ImageTypes[$extension] }
            ) }
        }
        if ($extension -eq ".pdf") {
            $data = [Convert]::ToBase64String([IO.File]::ReadAllBytes($path))
            return @{ __content = @(
                @{ type = "text"; text = $header },
                @{ type = "resource"; resource = @{ uri = "file:///" + $name; mimeType = "application/pdf"; blob = $data } }
            ) }
        }
        $text = $null
        if ($extension -eq ".docx") { $text = Read-Docx $path }
        elseif ($extension -eq ".xlsx") { $text = Read-Xlsx $path }
        elseif ($TextTypes -contains $extension) {
            $text = [IO.File]::ReadAllText($path)
            if ($extension -in @(".html", ".htm")) { $text = ConvertFrom-Markup $text "p" }
        }
        if ($null -eq $text) {
            throw "$name cannot be read here (only PDF, Word .docx, Excel .xlsx, pictures and text files). " +
                  "Save it with outlook_save_attachments and open it yourself."
        }
        if ($text.Length -gt $MaxTextChars) { $text = $text.Substring(0, $MaxTextChars) + "`n[cut at $MaxTextChars characters]" }
        return @{ __content = @(@{ type = "text"; text = "$header`n`n$text" }) }
    } finally {
        Remove-Item -LiteralPath $directory -Recurse -Force -ErrorAction SilentlyContinue
    }
}
