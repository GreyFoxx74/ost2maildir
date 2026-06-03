# ost2maildir

**Convert Microsoft Outlook OST files to Dovecot-compatible Maildir++ format.**

If you have an OST file from a decommissioned Exchange account, a corporate mailbox migration, or an orphaned Outlook profile, this tool extracts everything — email, attachments, calendar events, and contacts — into standard files that any IMAP server can import directly.

Tested against a real 3.5 GB corporate Exchange mailbox: **45,107 messages, 4,502 calendar items, 25 contacts — 0 errors.**

---

## Features

- **Complete message fidelity** — preserves all headers, HTML/plain bodies, and attachments
- **Full folder hierarchy** — Inbox sub-folders, Sent, Drafts, Trash, custom folders all mapped to Maildir++ convention
- **Attachment extraction** — inline images, file attachments, and TNEF-encoded (`winmail.dat`) Exchange messages decoded automatically
- **Calendar items** — `IPM.Appointment` events converted to `.ics` attachments, importable into any CalDAV server
- **Contacts** — `IPM.Contact` entries converted to `.vcf` vCard files
- **Read/replied/flagged state** — Outlook message flags mapped to standard Maildir `S`, `R`, `F`, `D`, `P` flags
- **Deduplication** — skips messages with duplicate `Message-ID` headers; run is resumable if interrupted
- **Handles compressible encryption** — the default OST encryption mode used by Exchange Cached Mode is decoded transparently
- **Graceful error handling** — corrupted OST nodes are logged and skipped; the rest of the mailbox continues

---

## Requirements

- Python 3.10 or newer
- Linux or WSL2 (Ubuntu 22.04+ recommended)
- ~1–1.5× the OST file size in free disk space for output

---

## Installation

```bash
git clone https://github.com/GreyFoxx74/ost2maildir.git
cd ost2maildir

python3 -m venv .venv
source .venv/bin/activate

pip install -e .
```

This installs the `ost2maildir` command and all dependencies, including `libpff-python` (the C library that reads OST/PST binary format).

---

## Quick Start

```bash
source .venv/bin/activate

# 1. Dry-run first — counts everything without writing any files
ost2maildir --dry-run mailbox.ost ./Maildir

# 2. Convert
ost2maildir mailbox.ost ./Maildir

# 3. Import into Dovecot
doveadm import -u username@example.com maildir:/path/to/Maildir "" "*"
```

### Example dry-run output

```
ost2maildir v1.0.0  |  2026-06-02 15:21
Input : mailbox.ost
Output: ./Maildir
Encryption: compressible

Dry-run mode — counting items...

  Email messages :   40,713
  Calendar items :    4,502
  Contacts       :       25
  Tasks/other    :        6
  Parse errors   :        0
  TOTAL          :   45,248

  Estimated output: ~0.6 GB
```

---

## Output Structure

The output is a standard Maildir++ directory tree ready for Dovecot:

```
Maildir/
├── cur/                    ← INBOX messages
├── new/
├── tmp/
├── .Sent/
├── .Drafts/
├── .Trash/
├── .Junk/
├── .Archive/
├── .Calendar/              ← appointments wrapped as .ics messages
├── .Contacts/              ← contacts wrapped as .vcf messages
├── .CustomFolder/          ← any Outlook folder you created
├── .Parent.Child/          ← nested folders use dot separator
├── subscriptions           ← Dovecot folder subscription list
└── .ost2maildir_seen       ← dedup state (delete after import)
```

---

## All Options

```
ost2maildir [OPTIONS] INPUT_OST OUTPUT_MAILDIR

Options:
  --dry-run          Count items without writing output
  --no-dedup         Disable Message-ID deduplication
  --skip-contacts    Skip contact conversion
  --skip-calendar    Skip calendar conversion
  --no-tnef          Keep winmail.dat without decoding
  --encoding TEXT    Fallback charset for ANSI bodies  [default: cp1252]
  --log-file PATH    Write full log to file
  --hostname TEXT    Override hostname in generated Message-IDs
  -v, --verbose      Debug-level logging
  --version          Show version and exit
  --help             Show this message and exit
```

---

## Importing into Dovecot

After conversion, use `doveadm` to import into a live Dovecot account:

```bash
# Import all folders
doveadm import -u username@example.com maildir:/path/to/Maildir "" "*"

# Rebuild indexes
doveadm index -u username@example.com "*"
```

Or copy directly if the Maildir is already on the mail server:

```bash
rsync -av --chown=vmail:vmail ./Maildir/ /var/mail/vhosts/example.com/username/
```

---

## Folder Mapping

Standard Outlook folders are mapped to Dovecot-conventional names automatically:

| Outlook | Maildir++ |
|---|---|
| Inbox | _(root)_ |
| Sent Items | `.Sent` |
| Deleted Items | `.Trash` |
| Drafts | `.Drafts` |
| Junk E-mail | `.Junk` |
| Archive | `.Archive` |
| Inbox/AnySubFolder | `.AnySubFolder` |
| Work/Projects | `.Work.Projects` |

Internal Exchange folders (`Recipient Cache`, `Finder`, `Drizzle`, etc.) are excluded automatically.

---

## Known Limitations

| Limitation | Notes |
|---|---|
| **High encryption OST** | OST files with `0x02` high encryption cannot be read by libpff. Export to PST via Outlook first, then convert the PST. |
| **Exchange-internal messages** | Messages that never left Exchange (no SMTP relay) have no `Received:` chain. Headers are reconstructed from MAPI properties. |
| **Calendar recurrence** | Simple daily/weekly/monthly rules export correctly. Complex patterns (e.g., "last business day of month") export as single non-recurring events. |
| **Calendar location / attendees** | Not currently extracted. |
| **S/MIME and PGP** | Encrypted message bodies are preserved as-is; not decrypted. |

See [USAGE.md](USAGE.md) for the full guide, including troubleshooting.

---

## How It Works

The tool reads the OST binary format using [libpff](https://github.com/libyal/libpff), which implements the [MS-PST] Open Specification published by Microsoft. For each MAPI message object it:

1. Reads MAPI properties (sender, recipients, dates, flags, body, attachments)
2. Constructs a valid RFC 2822 MIME message
3. Decodes any TNEF blobs (`winmail.dat`) to recover hidden attachments
4. Maps the Outlook folder path to a Maildir++ directory name
5. Writes the file atomically (`tmp/` → `cur/`) with the correct Maildir flag suffix

Calendar and contact items follow the same pipeline but are wrapped as RFC 2822 messages carrying `.ics` or `.vcf` attachments, keeping them importable via standard IMAP clients.

---

## License

MIT — see [LICENSE](LICENSE)

Dependencies include LGPL-licensed libraries (libpff, icalendar, chardet) — see [NOTICE](NOTICE) for attribution.
