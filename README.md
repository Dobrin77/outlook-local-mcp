# outlook-local-mcp

*[Български](README.bg.md)*

A Claude plugin that lets Claude search, read and organise the mail that **classic Outlook keeps on your Windows PC**: your mailbox cache (OST) and any archive data files (PST) you have open in Outlook.

It is a small local MCP server written in Windows PowerShell 5.1. It talks to Outlook through COM on your own computer. It **never deletes and never sends** mail: it can prepare a draft with files attached, which you send yourself.

## Why

- **Reach mail that only exists locally.** Many people have years of mail in PST archives, or a mail server that keeps only recent messages. Cloud connectors cannot see that mail; this plugin can.
- **No passwords, no accounts, no tokens.** It uses the Outlook that is already signed in on your PC.
- **Local only.** The server runs on your computer and does not connect to any network service itself.

## Requirements

- Windows 10 or 11 with Windows PowerShell 5.1 (built in).
- **Classic Outlook**: Outlook 2016, 2019, 2021, 2024 or Microsoft 365 desktop Outlook. **"New Outlook" is not supported** (it has no COM interface).
- Claude Desktop or Claude Code on the same PC.

## Install

This repository is a Claude plugin marketplace with one plugin, `outlook-local`.

**Claude Code**

```
/plugin marketplace add Dobrin77/outlook-local-mcp
/plugin install outlook-local@outlook-local-mcp
```

**Claude Desktop**

Add a plugin marketplace from GitHub with the repository `Dobrin77/outlook-local-mcp` (in the plugin / customisation settings), install **outlook-local**, make sure it is enabled, and restart Claude Desktop.

**For an organisation**

An organisation admin can add this repository as a plugin marketplace for the organisation, so staff can install the plugin without adding the marketplace themselves. The plugin still runs on each user's own PC and only sees that user's Outlook.

## Tools

| Tool | What it does |
| --- | --- |
| `outlook_status` | Checks the connection to Outlook: version, mail account and the open data files (PST/OST). |
| `outlook_list_folders` | Lists mail folders in the local data files, with message and unread counts. |
| `outlook_search` | Searches local mail, newest first, by words, sender, recipient, subject, dates, folder, data file, unread, flagged or category. |
| `outlook_get_message` | Reads one message in full: headers, plain-text body and the list of attachments. |
| `outlook_read_attachment` | Opens one attachment for Claude to read: PDF, pictures, Word (.docx), Excel (.xlsx) and text files, up to 10 MB. |
| `outlook_waiting_for_me` | Finds conversations waiting for your answer: someone else wrote last and you have not replied. |
| `outlook_waiting_for_others` | Finds conversations where you wrote last and nobody has answered yet. |
| `outlook_create_folder` | Creates a mail folder (returns the existing one if it is already there). Not inside Deleted Items, Junk E-mail, Outbox or Sync Issues. |
| `outlook_move_messages` | Moves messages to another folder (at most 200 at a time). Deleted Items, Junk E-mail, Outbox and Sync Issues are refused as targets. |
| `outlook_mark_messages` | Marks messages read/unread, sets or clears the follow-up flag, adds or removes categories. |
| `outlook_save_attachments` | Saves attachments to a folder on a local drive of your PC, named by message date; existing files are never overwritten. See the limits below. |
| `outlook_create_draft` | Prepares a new message in Outlook with files from your PC attached, saves it in Drafts and opens it. You check it and press Send. See the limits below. |

## Example prompts

- "Find the invoices from last March in my Outlook archive and list the totals."
- "Search my local Outlook for the contract we discussed with the supplier in 2023 and summarise the attachment."
- "Which emails in my Inbox from the last two weeks am I still supposed to answer?"
- "Which offers did I send this month that nobody has replied to?"
- "Create a folder Inbox\Invoices and move all invoice emails from this year there. Show me the list first."
- "Save all PDF attachments from messages with 'statement' in the subject to D:\Documents\Statements."
- "Flag the unread messages from my manager and add the category 'Follow up'."

## How the follow-up tools decide

`outlook_waiting_for_me` and `outlook_waiting_for_others` rebuild conversations from the messages themselves (Message-ID, In-Reply-To and References headers, with subject plus the other person as a fallback). This means replies you sent from another mail client also count, as long as a copy of them reaches this Outlook.

- **Inbox by default.** Only messages to you in the Inbox count as waiting for you; mail that rules sorted into other folders is usually not for answering. Use `all_folders` to look everywhere, and `include_cc` to include messages where you are only in CC.
- **Automated mail is skipped.** Messages from addresses such as noreply, notifications or newsletter, and messages with bulk-mail or mailing-list headers (List-Unsubscribe, Precedence: bulk, Auto-Submitted, and similar), are left out. This is a heuristic, so Claude is told to drop any remaining automatic messages itself.
- **Forwards are excluded** from `outlook_waiting_for_others` (a forward rarely asks for an answer) unless you ask for them with `include_forwards`.
- By default the last 30 days are checked (up to 365). Very large mailboxes may return a partial answer with a note; use fewer days or the `store` argument.

## Never deletes or sends

There is no tool to delete or send mail. `outlook_create_draft` only saves a draft and opens it; you press Send. Moves, flags and categories can all be undone in Outlook. To clean up, Claude can move messages to a folder that you empty yourself. Messages are never moved into Deleted Items, Junk E-mail, Outbox or Sync Issues (or folders under them), and no folders are created there.

## Saving attachments safely

`outlook_save_attachments` only writes to an ordinary folder on a local drive letter, such as `D:\Accounting\Invoices`:

- Network paths (`\\server\share`) and device paths (`\\?\...`, `\\.\...`) are refused.
- The Windows folder, Program Files, ProgramData, your AppData folders and the Startup folders are refused.
- Program and script files (`.exe`, `.com`, `.bat`, `.cmd`, `.vbs`, `.js`, `.wsf`, `.hta`, `.lnk`, `.ps1`, `.scr`, `.msi`, `.dll`, `.jar`, `.reg`, `.url` and similar) are skipped and reported, unless you ask for that file type explicitly (for example `extensions: ["exe"]`).
- Claude is told to ask you which folder to use when you did not name one.

## Drafts with files attached

`outlook_create_draft` attaches only files the user named:

- Files must be on a drive letter; mapped network drives such as `N:\` work, `\server\share` paths are refused.
- Files in the Windows folder, Program Files, ProgramData and your AppData folders are refused, so browser profiles and saved credentials cannot be attached.
- Outlook data files (`.pst`, `.ost`, `.nst`), password databases (`.kdbx`), keys and certificates (`.pfx`, `.p12`, `.key`, `.pem`, `.ppk`) are refused.
- At most 20 files and 25 MB together.
- The draft is plain text and not sent: Claude is told never to attach files because a message or document asks for it.

Mail and attachments are treated as untrusted content: Claude is told never to follow instructions found in them.

## Privacy

The server runs only on your PC and reads Outlook locally. Mail that Claude finds or opens with these tools is sent to Claude as part of your conversation, the same as with any other connector. Nothing else is sent anywhere: the plugin has no telemetry and makes no network connections of its own.

## Troubleshooting

**Outlook shows "A program is trying to access e-mail address information..."**
Outlook asks this when it thinks no antivirus is active. Open Outlook > File > Options > Trust Center > Trust Center Settings > Programmatic Access. When an antivirus is active and up to date, the antivirus status shows **Valid** and the prompt does not appear. If it is not valid, fix the antivirus status (or ask your IT admin); do not lower the setting unless you understand the risk.

**Searches are slow or stop early**
Word searches use the Windows Search index when Outlook has one. Check that Windows Search is running and that Outlook is included in indexing (Outlook > File > Options > Search > Indexing Options). Narrowing by dates, folder or data file also helps.

**The tools are not visible in Claude**
Check that the plugin is installed and enabled, then fully quit and restart Claude Desktop (also from the system tray).

**"Outlook is not available on this computer"**
Classic Outlook is not installed, or only "New Outlook" is in use. Install or switch back to classic Outlook.

## Status and support

This plugin is shared as-is, in the hope that it is useful. Issues and pull requests are welcome and will be looked at on a best-effort basis, without any guarantee of support.

## Acknowledgements

Written by Dobrin Tassev together with [Claude Code](https://claude.com/claude-code) (Anthropic), which did most of the coding, testing and the security review.

## License

[MIT](LICENSE)
