# Security Policy

## Supported versions

Only the latest version of the plugin (see `plugins/outlook-local/.claude-plugin/plugin.json`)
receives security fixes.

## Reporting a vulnerability

Please **do not** open a public issue for security problems.

Report privately through GitHub: open the **Security** tab of this repository
and click **Report a vulnerability**. Include what you found, how to reproduce
it, and the Outlook / Windows / PowerShell versions involved.

What to expect:

- an acknowledgement within 7 days;
- a fix or a decision within 30 days for confirmed issues;
- credit in the fix commit, if you want it.

This is a small project maintained in spare time, so timelines are best effort.

## Scope notes

The plugin is designed never to send or delete mail and to save attachments
only to local folders. Anything that breaks those guarantees, writes files
outside the chosen folder, or lets the content of an e-mail make the plugin
act on its own is treated as critical.
