# Security Policy

## Supported versions

Security fixes are provided for the latest release of Native PowerPoint Doc Editor.

## Reporting a vulnerability

Please report security issues privately instead of opening a public issue.

1. Email or message the project maintainers with a clear description of the issue.
2. Include reproduction steps and impact when possible.
3. Allow reasonable time for a fix before public disclosure.

We will acknowledge valid reports and coordinate a fix and disclosure timeline.

## Backups and incident response

This plugin edits vault-local DOCX/PPTX files in place. Keep your own backups of important documents (Obsidian Sync, Time Machine, git, or copies outside the vault). The plugin may write recovery copies when a save fails during close; treat those as last-resort, not a backup plan.

If you discover a security or data-loss incident (unexpected remote access, corrupted saves, leaked secrets in a shared vault):

1. Stop using the affected build and isolate the vault/copy if needed.
2. Report privately using the steps above.
3. Restore from a known-good backup before continuing edit work.
4. Rotate any credentials that may have been exposed in document content or plugin settings.
