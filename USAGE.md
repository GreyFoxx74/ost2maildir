# ost2maildir — Usage Guide

Convert a Microsoft Outlook OST file to Dovecot-compatible Maildir++ format for direct IMAP server import.

---

## Table of Contents

1. [Installation](#installation)
2. [Quick Start](#quick-start)
3. [Command Reference](#command-reference)
4. [Workflow: Step by Step](#workflow-step-by-step)
5. [Understanding the Output](#understanding-the-output)
6. [Importing into Dovecot](#importing-into-dovecot)
7. [Folder Mapping Reference](#folder-mapping-reference)
8. [Flags and Read State](#flags-and-read-state)
9. [Calendar and Contacts](#calendar-and-contacts)
10. [Troubleshooting](#troubleshooting)
11. [Known Limitations](#known-limitations)

---

## Installation

### Prerequisites

- Python 3.10 or newer
- Linux or WSL2 (tested on Ubuntu 22.04 / WSL2 with Python 3.12)

### Setup

```bash
# Clone or copy the project
cd /path/to/ost2maildir

# Create a virtual environment
python3 -m venv .venv
source .venv/bin/activate

# Install all dependencies (including libpff-python wheel)
pip install -e .
```

This installs the `ost2maildir` command into the virtual environment. All dependencies are self-contained.

### Dependency Notes

| Package | Purpose |
|---|---|
| `libpff-python` | Reads OST/PST binary format via libpff C library |
| `tnefparse` | Decodes `winmail.dat` TNEF attachments from Exchange RTF messages |
| `striprtf` | Converts RTF message bodies to plain text when no HTML version exists |
| `icalendar` | Generates `.ics` calendar files from Outlook appointment items |
| `vobject` | Generates `.vcf` vCard files from Outlook contact items |
| `click` | CLI framework |
| `tqdm` | Progress bar |
| `chardet` | Fallback charset detection for legacy ANSI message bodies |

---

## Quick Start

```bash
source .venv/bin/activate

# 1. Always run a dry-run first to assess the file
ost2maildir --dry-run mailbox.ost ./output_maildir

# 2. Convert with a log file
ost2maildir --log-file convert.log mailbox.ost ./output_maildir

# 3. Check the output
ls ./output_maildir/
find ./output_maildir -path "*/cur/*" -type f | wc -l
```

---

## Command Reference

```
ost2maildir [OPTIONS] INPUT_OST OUTPUT_MAILDIR
```

### Arguments

| Argument | Description |
|---|---|
| `INPUT_OST` | Path to the `.ost` or `.pst` file to convert. |
| `OUTPUT_MAILDIR` | Output directory. Created if it does not exist. This becomes the `~/Maildir/` root on the IMAP server. |

### Options

| Option | Default | Description |
|---|---|---|
| `--dry-run` | off | Scan and count all items without writing any output. Shows item counts and estimated output size. Run this first. |
| `--no-dedup` | off | Disable Message-ID deduplication. By default, messages with a duplicate `Message-ID` are written only once. |
| `--skip-contacts` | off | Skip conversion of `IPM.Contact` items. No `.Contacts` folder will be created. |
| `--skip-calendar` | off | Skip conversion of `IPM.Appointment` items. No `.Calendar` folder will be created. |
| `--no-tnef` | off | Preserve `winmail.dat` attachments as-is without decoding. By default TNEF blobs are decoded to expose hidden attachments. |
| `--encoding TEXT` | `cp1252` | Fallback character encoding for message bodies that cannot be decoded as UTF-8. Common alternatives: `latin-1`, `iso-8859-1`. |
| `--log-file PATH` | none | Write full conversion log to this file in addition to stderr output. |
| `--hostname TEXT` | system hostname | Override the hostname used in synthesized `Message-ID` values and Maildir filenames. |
| `-v, --verbose` | off | Enable debug-level logging. Logs every skipped message and MAPI property access. |
| `--version` | — | Show version and exit. |
| `--help` | — | Show help and exit. |

---

## Workflow: Step by Step

### Step 1 — Assess the OST file

Always start with a dry-run. It reads the entire file structure without writing anything:

```bash
ost2maildir --dry-run mailbox.ost /tmp/unused
```

Expected output:
```
ost2maildir v1.0.0  |  2026-06-02 15:21
Input : mailbox.ost
Output: /tmp/unused
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

**What to look for:**
- `Encryption: compressible` — this is normal and the tool handles it automatically.
- `Encryption: high` — the file cannot be converted directly. See [Troubleshooting](#troubleshooting).
- `Parse errors: 0` — ideal. Any non-zero count means some OST nodes are corrupted; they are skipped gracefully.
- Estimated output size — ensure you have at least this much free disk space, ideally 1.5× more.

### Step 2 — Run the conversion

```bash
ost2maildir --log-file convert.log mailbox.ost ./Maildir
```

Progress is shown as a live counter. The description shows the current folder being processed.

On a modern SSD, expect roughly **80–200 messages per second**. A 45,000-message mailbox typically takes 5–15 minutes depending on attachment sizes.

### Step 3 — Verify the output

```bash
# Count messages per folder
find ./Maildir -maxdepth 1 -type d | sort | while read dir; do
    count=$(find "$dir/cur" -type f 2>/dev/null | wc -l)
    [ "$count" -gt 0 ] && echo "  $(basename "$dir"): $count"
done

# Total message count
find ./Maildir -path "*/cur/*" -type f | wc -l

# Check for any conversion errors
grep -i "error\|warn" convert.log
```

### Step 4 — Import into Dovecot

See the [Importing into Dovecot](#importing-into-dovecot) section.

---

## Understanding the Output

### Directory Structure

The output is a **Maildir++** directory tree. After a successful conversion it looks like:

```
./Maildir/
├── cur/                    ← INBOX messages (read or flagged)
├── new/                    ← INBOX messages (unread, delivered fresh)
├── tmp/                    ← Temporary (always empty after conversion)
├── .Sent/
│   ├── cur/
│   ├── new/
│   └── tmp/
├── .Drafts/
├── .Trash/
├── .Calendar/              ← Appointments as .ics-wrapped messages
├── .Contacts/              ← Contacts as .vcf-wrapped messages
├── .Bulletins/             ← Example: custom Inbox subfolder
├── .SyncIssues.Conflicts/  ← Example: nested folder (SyncIssues/Conflicts)
├── subscriptions           ← Dovecot folder subscription list
└── .ost2maildir_seen       ← Deduplication state (safe to delete after import)
```

### Message Filenames

Each message file is named following the Maildir standard:

```
1718000000.42_12345.hostname:2,RS
│           │  │     │        │ └─ Flags: R=Replied, S=Seen
│           │  │     │        └─── Info version separator
│           │  │     └─────────── Hostname
│           │  └───────────────── PID
│           └──────────────────── Sequence number
└──────────────────────────────── Unix timestamp (message delivery time)
```

### The `subscriptions` File

This file tells Dovecot which folders to show the user. It contains one folder name per line (without the leading dot). It is created automatically at the end of conversion.

```
Sent
Drafts
Trash
Calendar
Contacts
Bulletins
SyncIssues.Conflicts
```

### The `.ost2maildir_seen` File

This JSON file records every `Message-ID` that was written. It serves two purposes:

1. **Deduplication during a single run** — prevents the same message being written twice (common when messages are in both INBOX and a user sub-folder).
2. **Resumable runs** — if conversion is interrupted and re-run against the same output directory, already-converted messages are skipped.

You can safely delete this file after you have confirmed the import is complete.

---

## Importing into Dovecot

### Option 1 — Direct Copy (simplest)

If the Maildir output is already on the mail server, copy it to the user's mail location:

```bash
# Adjust path to match your Dovecot mail_location setting
rsync -av --chown=vmail:vmail ./Maildir/ /var/mail/vhosts/example.com/username/

# Rebuild Dovecot indexes
doveadm index -u username@example.com INBOX
doveadm index -u username@example.com "*"
```

### Option 2 — `doveadm import` (recommended)

`doveadm import` is the official way to import mail into a Dovecot account. It handles UID assignment, deduplication, and index building:

```bash
# Import all folders
doveadm import -u username@example.com maildir:/path/to/Maildir "" "*"

# Import only the INBOX
doveadm import -u username@example.com maildir:/path/to/Maildir "" "INBOX"

# Import with verbose output
doveadm -v import -u username@example.com maildir:/path/to/Maildir "" "*"
```

### Option 3 — `doveadm sync` (for replication setups)

```bash
doveadm sync -u username@example.com maildir:/path/to/Maildir
```

### Post-import checks

```bash
# List all folders visible to the user
doveadm mailbox list -u username@example.com

# Check message count in INBOX
doveadm stat -u username@example.com INBOX

# Check message count in all folders
doveadm mailbox status -u username@example.com "messages" "*"
```

---

## Folder Mapping Reference

The following standard Outlook folder names are automatically remapped to Dovecot-conventional names:

| Outlook Folder Name | Maildir++ Folder | Notes |
|---|---|---|
| `Inbox` | _(root)_ | INBOX maps to the Maildir root directory |
| `Sent Items` | `.Sent` | |
| `Sent` | `.Sent` | Some tenants use this name |
| `Deleted Items` | `.Trash` | |
| `Trash` | `.Trash` | |
| `Drafts` | `.Drafts` | |
| `Junk E-mail` | `.Junk` | |
| `Junk Email` | `.Junk` | |
| `Junk` | `.Junk` | |
| `Spam` | `.Junk` | |
| `Outbox` | `.Outbox` | |
| `Archive` | `.Archive` | |
| `Calendar` | `.Calendar` | Appointments written here as `.ics` messages |
| `Contacts` | `.Contacts` | Contacts written here as `.vcf` messages |

**Sub-folder naming:**

- Sub-folders of `Inbox` are promoted to top-level Maildir++ folders: `Inbox/Bulletins` → `.Bulletins`
- All other nested folders use the dot separator: `Work/Projects/Alpha` → `.Work.Projects.Alpha`
- Characters that are unsafe for Maildir++ or filesystems (`/`, `\`, `.`, `:`, `*`, `?`, `"`, `<`, `>`, `|`) are replaced with `_`

**Skipped folders** (internal Outlook/Exchange metadata, never contain user mail):

- `Common Views`, `Finder`, `Views`, `Shortcuts`, `Drizzle`, `Shared Data`
- `Recipient Cache`, `GAL Contacts`, `Organizational Contacts`
- `Skype for Business Contacts`, `PeopleCentricConversation Buddies`
- `Yammer Root`, `Social Activity Notifications`, `RSS Feeds`
- `EventCheckPoints`, `PersonMetadata`, `Conversation Action Settings`
- `Quick Step Settings`, `Sync Issues` sub-folders

---

## Flags and Read State

Outlook read/replied/flagged state is preserved in Maildir flags:

| Maildir Flag | Letter | Outlook State |
|---|---|---|
| Seen (read) | `S` | Message was opened/read in Outlook |
| Draft | `D` | Message in Drafts folder (never sent) |
| Replied | `R` | User clicked Reply or Reply All |
| Passed (forwarded) | `P` | User clicked Forward |
| Flagged | `F` | Follow-up flag set in Outlook |
| Trashed | `T` | _(not currently mapped)_ |

**Placement in `cur/` vs `new/`:**

All messages are written to `cur/`. If a message is unread, its filename will not contain the `S` flag (e.g., `:2,` with no flags). IMAP clients and Dovecot correctly interpret the absence of `S` as unread.

---

## Calendar and Contacts

### Calendar Items

Each Outlook appointment (`IPM.Appointment`) is converted to a standard RFC 2822 email message containing a `.ics` (iCalendar) attachment. The message subject is prefixed with `[Calendar]`.

These messages land in the `.Calendar` Maildir++ folder.

**What is preserved:**
- Event title (SUMMARY)
- Start and end date/time (DTSTART, DTEND)
- Event description/notes (DESCRIPTION)
- Privacy class: Public / Private / Confidential (CLASS)
- Importance level (PRIORITY)
- Unique identifier (UID, deterministically generated from OST node ID)

**What is not preserved:**
- Attendees and RSVP state
- Complex recurrence rules (simple daily/weekly/monthly patterns are converted; complex patterns generate a single non-recurring VEVENT with a warning)
- Location (stored in a named MAPI property that requires Named Property Map resolution)
- Reminders and alarms

**Importing calendar items** into a CalDAV server (e.g., Radicale, Nextcloud):
```bash
# Extract all .ics files from the Calendar folder
for f in ./Maildir/.Calendar/cur/*; do
    python3 -c "
import email, sys
msg = email.message_from_bytes(open('$f','rb').read())
for part in msg.walk():
    if part.get_content_type() == 'text/calendar':
        print(part.get_payload(decode=True).decode('utf-8', errors='replace'))
        break
"
done > all_events.ics
```

### Contacts

Each Outlook contact (`IPM.Contact`) is converted to an RFC 2822 message containing a `.vcf` (vCard 3.0) attachment. The message subject is prefixed with `[Contact]`.

These messages land in the `.Contacts` Maildir++ folder.

**What is preserved:**
- First name, last name, full name
- Company and job title
- Work, home, and mobile phone numbers
- Email address (SMTP addresses only; Exchange X.500 addresses are omitted)
- Notes/body text

**Note on Recipient Cache contacts:** Outlook automatically populates a `Recipient Cache` folder with auto-complete entries from sent mail. This folder is excluded because the entries are not real user-managed contacts.

**Importing contacts** into a CardDAV server:
```bash
# Extract all .vcf files from the Contacts folder
for f in ./Maildir/.Contacts/cur/*; do
    python3 -c "
import email
msg = email.message_from_bytes(open('$f','rb').read())
for part in msg.walk():
    if 'vcard' in part.get_content_type():
        print(part.get_payload(decode=True).decode('utf-8', errors='replace'))
        break
"
done > all_contacts.vcf
```

---

## Troubleshooting

### "High encryption" error

```
ERROR: OST file uses high encryption (0x02) — cannot be read by libpff.
```

**Cause:** Your Exchange administrator enabled strong encryption via Group Policy. libpff cannot decrypt this cipher without the original Exchange profile key.

**Fix:**
1. Open the OST in Outlook (requires the original Exchange/Microsoft 365 account to be active).
2. Go to **File → Open & Export → Import/Export**.
3. Select **Export to a file → Outlook Data File (.pst)**.
4. Export all folders.
5. Run `ost2maildir` on the resulting `.pst` file (identical format, different encryption flag).

---

### Duplicate messages not expected

**Symptom:** Messages from sub-folders like `ToACTION` or custom sorting folders appear only in INBOX, not in their original sub-folder.

**Cause:** This is expected behavior. When you move or copy a message to a sub-folder in Outlook using Cached Exchange Mode, the OST sometimes stores the message in both locations. The deduplicator (by `Message-ID`) writes it once, to whichever folder was processed first.

**Fix:** Use `--no-dedup` to write every occurrence of every message, preserving all folder placements at the cost of duplicates:

```bash
ost2maildir --no-dedup mailbox.ost ./Maildir
```

---

### Low message count vs. expected

**Symptom:** Dry-run shows fewer messages than expected.

**Likely causes:**
1. **Recipient Cache** — the `Recipient Cache` folder (often 500–2000 entries) is intentionally skipped. These are auto-complete history, not real emails.
2. **System folders** — `Finder`, `Drizzle`, `Shared Data`, and similar Outlook internal folders are skipped.
3. **Tasks/Notes** — `IPM.Task` and `IPM.StickyNote` items are out of scope and not counted as emails.

---

### Corrupted attachments or missing attachment data

**Symptom:** An attachment appears in the output but has 0 bytes, or `get_attachment()` raises an OSError.

**Cause:** Some OST nodes have corrupted local descriptor tables (often from incomplete sync operations). libpff cannot read these attachment data blocks.

**Behavior:** The tool logs a warning and skips the attachment. The message itself is still written with any other attachments intact.

**What you'll see in the log:**
```
WARNING ost2maildir.converter: Skipping attachment 2: libpff_local_descriptors_node...
```

---

### `winmail.dat` still present in output

**Symptom:** Some messages in the output still contain a `winmail.dat` attachment.

**Cause:** The TNEF decoder (`tnefparse`) failed to parse the specific TNEF blob. This can happen with malformed TNEF data.

**Behavior:** The tool falls back to preserving the `winmail.dat` as-is and logs a warning:
```
WARNING ost2maildir.tnef: TNEF decode failed; keeping winmail.dat as-is: ...
```

You can view these attachments with any TNEF viewer (e.g., `tnef` command-line tool, or online viewers).

---

### OSError during conversion

**Symptom:** Messages in a specific folder cause OSError exceptions.

**Cause:** The OST file has corrupted B-tree nodes for those messages (common after a failed sync or abrupt Outlook shutdown).

**Behavior:** The tool logs the error at WARNING level and skips the affected message. All other messages in the folder are still processed.

**To see which messages were skipped:**
```bash
ost2maildir --log-file convert.log -v mailbox.ost ./Maildir
grep "Skipping message" convert.log
```

---

### Resuming an interrupted conversion

If conversion is interrupted (Ctrl+C, power loss, etc.), simply re-run the exact same command. The `.ost2maildir_seen` file tracks which Message-IDs were already written. Messages already in the output will be skipped; only new messages will be written.

```bash
ost2maildir --log-file convert.log mailbox.ost ./Maildir
```

---

### Python `UnicodeDecodeError`

**Symptom:** A message body cannot be decoded.

**Cause:** The message body uses an ANSI code page that is not `cp1252` (the default fallback).

**Fix:** Try a different fallback encoding:
```bash
ost2maildir --encoding iso-8859-1 mailbox.ost ./Maildir
# or
ost2maildir --encoding latin-1 mailbox.ost ./Maildir
```

The tool also tries `utf-8`, `cp1252`, and `latin-1` in sequence before giving up, so this is rare.

---

## Known Limitations

| Limitation | Details |
|---|---|
| **High encryption** | OST files with `0x02` high encryption cannot be read. Convert via Outlook to PST first. |
| **X.500 email addresses** | Messages sent between Exchange users on the same server may lack SMTP `From:` addresses; the display name is preserved but the email address defaults to `unknown@ost2maildir.local`. |
| **No `Received:` headers** | Exchange-internal messages (never processed by SMTP) will not have `Received:`, `Return-Path:`, or `Authentication-Results:` headers in the output. |
| **Calendar recurrence** | Complex recurrence rules (e.g., "last business day of the month") are exported as non-recurring single events. Simple daily/weekly/monthly rules are exported correctly. |
| **Calendar location** | The appointment location field requires Named Property Map resolution which is not currently implemented. Location is omitted from `.ics` output. |
| **Calendar attendees** | Attendee lists and RSVP state are not currently extracted. |
| **S/MIME and PGP** | Encrypted message bodies are preserved as opaque MIME parts. The tool does not decrypt them. |
| **Embedded messages** | Messages attached as `.msg` files are attached as `application/octet-stream`. They are not recursively converted. |
| **Distribution lists** | `IPM.DistList` items are converted to vCard GROUP entries, but member X.500 addresses are dropped. |
| **Large dedup file** | The `.ost2maildir_seen` JSON file grows at ~60 bytes per message. For mailboxes over 500,000 messages, consider switching to a SQLite-backed dedup store. |

---

## Examples

**Convert a mailbox, skip calendar and contacts:**
```bash
ost2maildir --skip-calendar --skip-contacts mailbox.ost ./Maildir
```

**Convert only the Inbox and Sent Items (folder filter not yet implemented — use full conversion and selectively import with doveadm):**
```bash
ost2maildir mailbox.ost ./Maildir
doveadm import -u user@domain maildir:./Maildir INBOX INBOX
doveadm import -u user@domain maildir:./Maildir Sent Sent
```

**Convert a PST file (identical format):**
```bash
ost2maildir archive.pst ./Maildir
```

**Re-run after interruption (safe — skips already-written messages):**
```bash
ost2maildir --log-file convert.log mailbox.ost ./Maildir
```

**Debug a specific conversion problem:**
```bash
ost2maildir -v --log-file debug.log mailbox.ost ./Maildir
grep -i "error\|warn\|skip" debug.log | head -50
```

---

*ost2maildir v1.0.0 | Built and tested on WSL2 / Ubuntu with Python 3.12 and a 3.5 GB corporate Exchange OST file (45,107 messages, 4,502 calendar items, 25 contacts — 0 errors).*
