# OST2MAILDIR — Technical Design and Implementation Guide

> Comprehensive specification for a Python tool that converts Microsoft Outlook OST files
> to Dovecot-compatible Maildir++ format, preserving messages, attachments, calendar
> items, and contacts for direct IMAP server import.

---

## Table of Contents

1. [Overview and Goals](#1-overview-and-goals)
2. [Format Analysis: OST / PFF Internals](#2-format-analysis-ost--pff-internals)
3. [Format Analysis: Maildir / Maildir++ Internals](#3-format-analysis-maildir--maildir-internals)
4. [Technical Challenges](#4-technical-challenges)
5. [Architecture](#5-architecture)
6. [Dependency Stack](#6-dependency-stack)
7. [Property Mappings](#7-property-mappings)
8. [Implementation Plan (Phased)](#8-implementation-plan-phased)
9. [CLI Design and Usage](#9-cli-design-and-usage)
10. [Configuration Reference](#10-configuration-reference)
11. [Caveats and Known Limitations](#11-caveats-and-known-limitations)
12. [Testing Strategy](#12-testing-strategy)
13. [Performance Considerations](#13-performance-considerations)

---

## 1. Overview and Goals

### 1.1 Problem Statement

Microsoft Outlook stores its offline cache in an OST (Offline Storage Table) file —
a binary proprietary format tied to an Exchange/Microsoft 365 account profile. When
migrating email to a self-hosted IMAP server, or recovering mail from an orphaned
mailbox, the OST must be converted to a format the IMAP server understands natively.

Maildir is the standard file-per-message format supported by virtually every modern
IMAP server (Dovecot, Cyrus, Postfix, Mutt, etc.). Maildir++ extends Maildir with
a dot-prefix naming convention for nested folder hierarchies.

### 1.2 Scope

The tool will:
- Read an OST file of any size (tested target: 2–20 GB, corporate Exchange)
- Extract all email messages with full headers, bodies, and attachments
- Extract calendar items as iCal (`.ics`) files wrapped in RFC 2822 messages
- Extract contacts as vCard (`.vcf`) files wrapped in RFC 2822 messages
- Decode TNEF-encoded (`winmail.dat`) messages, restoring hidden attachments
- Produce a Maildir++ directory tree suitable for direct Dovecot import
- Preserve the original Outlook folder hierarchy
- Deduplicate messages by Message-ID
- Map Outlook read/flagged/replied state to standard Maildir flags
- Auto-detect OST encryption type and fail gracefully on unsupported ciphers

### 1.3 Out of Scope

- OST files using **high encryption** (see §11.1) — requires Outlook itself
- Outlook Tasks and Sticky Notes (no standard open format; low IMAP utility)
- Recovery of permanently deleted messages outside the OST's message store
- Bi-directional sync; this is one-time migration only
- Windows-only Outlook APIs (no COM automation dependency)

---

## 2. Format Analysis: OST / PFF Internals

### 2.1 File Format Foundation

OST and PST files share the same on-disk format, formally documented by Microsoft
as the **[MS-PST] Personal Storage Table** open specification
(https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-pst). Microsoft
published this under an irrevocable Open Specification Promise including free patent
licensing.

The binary container is called **PFF (Personal Folder File)**. Internally it is a
page-based B-tree database with two flavors:

| Attribute | ANSI (32-bit) | Unicode (64-bit) |
|---|---|---|
| Max file size | 2 GB | Effectively unlimited |
| Encoding | Windows code pages | UTF-16LE |
| Outlook version | Outlook 97–2002 | Outlook 2003+ |
| Common for OST | Rare (legacy) | Standard |

Modern corporate OST files are almost always Unicode format. The tool must handle both.

### 2.2 Internal Structure

The PFF format is a layered data store:

```
PFF File
├── File Header          (magic, format version, encryption type, root pointers)
├── Node B-tree (NBT)    (maps 64-bit node IDs → file offsets + sizes)
├── Block B-tree (BBT)   (maps 64-bit block IDs → data block locations)
├── Message Store        (root MAPI object, holds folder tree pointer)
└── Folder Tree
    ├── Root Folder      ("Top of Personal Folders" or "Mailbox - User")
    │   ├── Inbox
    │   ├── Sent Items
    │   ├── Deleted Items
    │   ├── Drafts
    │   ├── ...
    │   └── User Folders/
    │       └── Sub-folders/
    └── (Orphan nodes recovered separately)
```

Each **folder** is a MAPI object containing:
- A hierarchy table (sub-folder list)
- A contents table (message list)
- A FAI (Folder Associated Information) table (rules, views, calendar metadata)

Each **message** is a MAPI object containing:
- A property set (MAPI properties: subject, sender, dates, flags, body, etc.)
- A recipient table (To/CC/BCC addresses with display names and SMTP addresses)
- An attachment table (zero or more attachment objects)

Each **attachment** is a MAPI object with its own property set, including the raw
binary data stored in a data block or embedded as another message object (nested).

### 2.3 Named Property Map (NPM)

Not all MAPI properties have fixed IDs. Custom properties (used by Outlook
extensions, Exchange features, and some calendar/contact fields) are stored in the
Named Property Mapping table, which maps a GUID+name or GUID+ID pair to a local
property tag in the range 0x8000–0xFFFF. The tool must read this mapping table to
decode extended properties reliably.

### 2.4 Encryption Modes

Every PFF file has an encryption flag in the header. Three values are defined:

| Value | Name | Mechanism | libpff readable? |
|---|---|---|---|
| `0x00` | No encryption | None | Yes |
| `0x01` | Compressible | Fixed 256-byte substitution table; trivially reversible | Yes (automatic) |
| `0x02` | High encryption | Multi-rotor substitution cipher (Enigma-like) | **No** |

**Critical caveat:** The default encryption for OST files created by Exchange
Cached Mode is **compressible encryption (0x01)**. This is *not* real security; it
is data obfuscation to prevent accidental corruption by non-Outlook tools. libpff
handles it transparently.

**High encryption (0x02)** is rare but real. Some Exchange administrators enforce it
via Group Policy (`HKLM\SOFTWARE\Policies\Microsoft\Office\...`). If encountered,
the tool must detect it and abort with a meaningful error message rather than
producing corrupted output.

Detection code:
```python
import pypff

f = pypff.file()
f.open(path)
# libpff raises pypff.error for high-encrypted files
# The error message will contain "encryption"
```

### 2.5 OST-Specific Differences from PST

| Aspect | PST | OST |
|---|---|---|
| Primary use | Export/archive | Cached Exchange Mode |
| Independence | Standalone | Tied to Exchange profile GUID |
| MAPI Entry ID | Standard | Includes Exchange GUID prefix |
| Internet headers | Usually present | Often absent (internal Exchange messages) |
| Deleted items | Standard hierarchy | May include "Recoverable Items" dumpster |

The most important practical difference: **Exchange-internal messages** (sent between
users on the same Exchange server) were never processed by SMTP, so they have no
`PR_TRANSPORT_MESSAGE_HEADERS` property. Their From/To/Date must be reconstructed
entirely from MAPI properties.

### 2.6 Message Classes (MAPI Item Types)

Every MAPI object has a `PR_MESSAGE_CLASS` property (IPM = Interpersonal Message):

| Message Class | Type | Handling |
|---|---|---|
| `IPM.Note` | Email | Primary conversion target |
| `IPM.Appointment` | Calendar | Convert to .ics |
| `IPM.Contact` | Contact | Convert to .vcf |
| `IPM.Task` | Task | Out of scope |
| `IPM.StickyNote` | Note | Out of scope |
| `IPM.Schedule.Meeting.*` | Meeting request | Treat as email + .ics attachment |
| `IPM.DistList` | Distribution list | Treat as .vcf group card |
| `Report.*` | NDR/delivery report | Treat as email |
| `IPM.Note.Rules.*` | Out-of-office / rule | Treat as email |

### 2.7 pypff Python API Reference

The `pypff` module (from libpff Python bindings) exposes:

```python
import pypff

# Open file
pff = pypff.file()
pff.open("mailbox.ost")

# Navigate folder tree
root = pff.get_root_folder()
# root.number_of_sub_folders
# root.sub_folders  (iterable)
# root.get_sub_folder(index) -> pypff.folder

# Folder properties
folder.name                          # str: folder display name
folder.number_of_sub_folders         # int
folder.number_of_messages            # int
folder.sub_folders                   # iterable of pypff.folder
folder.messages                      # iterable of pypff.message

# Message properties (directly exposed)
msg.subject                          # str or None: PR_SUBJECT
msg.sender_name                      # str or None: PR_SENDER_NAME
msg.conversation_topic               # str or None
msg.transport_headers                # str or None: full SMTP header block
msg.plain_text_body                  # str or None: PR_BODY
msg.html_body                        # bytes or None: PR_HTML
msg.rtf_body                         # bytes or None: PR_RTF_COMPRESSED (decompressed)
msg.client_submit_time               # datetime or None: PR_CLIENT_SUBMIT_TIME
msg.delivery_time                    # datetime or None: PR_MESSAGE_DELIVERY_TIME
msg.creation_time                    # datetime or None: PR_CREATION_TIME
msg.flags                            # int: PR_MESSAGE_FLAGS bitmask
msg.identifier                       # int: internal node ID (useful for filename)
msg.number_of_attachments            # int
msg.get_attachment(index)            # pypff.attachment

# Full MAPI property access (for non-exposed properties like To, CC, Message-ID)
for record_set in msg.record_sets:
    for entry in record_set.entries:
        entry.entry_type   # int: MAPI property ID (e.g., 0x0E04 for PR_DISPLAY_TO)
        entry.data_as_string  # decoded string value (if applicable)
        entry.data         # raw bytes

# Attachment properties
att.name                             # str: display name / filename
att.size                             # int: byte size
att.read_buffer(chunk_size)          # streaming read
att.data                             # bytes: full attachment data (avoid for large files)
att.record_sets                      # MAPI properties (for long filename, mime type, etc.)
```

---

## 3. Format Analysis: Maildir / Maildir++ Internals

### 3.1 Maildir Base Format

Maildir was designed by D.J. Bernstein. Each mailbox is a directory containing three
subdirectories:

```
maildir/
├── tmp/    # Temporary files during delivery (in-flight only)
├── new/    # Newly delivered, unread messages
└── cur/    # Messages the mail client has seen (read or unread)
```

**Delivery protocol**: messages are created in `tmp/`, then atomically hard-linked to
`new/`. When a client marks a message as seen, it moves it to `cur/` and appends the
flags suffix to the filename.

For import purposes (not live delivery), messages go directly into `cur/` with the
appropriate flags set. This is standard practice for IMAP migration.

### 3.2 Maildir Filename Convention

Messages in `cur/` follow this filename format (info flags version 2):

```
<timestamp>.<pid>.<hostname>[,S=<filesize>]:2,<flags>
```

| Component | Description |
|---|---|
| `timestamp` | Unix epoch seconds (integer, the message delivery time) |
| `pid` | Process ID or any unique integer |
| `hostname` | Hostname (with `:` and `/` replaced by `\057` and `\072`) |
| `,S=<n>` | File size in bytes (optional; used by Dovecot for quota) |
| `:2,` | Fixed info-flags version separator |
| `flags` | Alphabetically sorted single-character flags |

Example:
```
1718000000.1234.mailserver,S=4096:2,RS
```
(timestamp=1718000000, pid=1234, hostname=mailserver, size=4096, flags=Read+Seen)

**Flags (standard):**

| Letter | Name | IMAP equivalent | Outlook equivalent |
|---|---|---|---|
| `D` | Draft | `\Draft` | Drafts folder item |
| `F` | Flagged | `\Flagged` | Follow-up flag |
| `P` | Passed | (forwarded) | Forwarded message |
| `R` | Replied | `\Answered` | Replied indicator |
| `S` | Seen | `\Seen` | Read |
| `T` | Trashed | `\Deleted` | Deleted flag |

Flags **must be in alphabetical order** in the filename or some IMAP servers will
reject the message or miscategorize it.

### 3.3 Maildir++ Nested Folder Convention

Standard Maildir supports only one folder. Maildir++ (an extension by the Courier
Mail Server project, adopted universally) adds hierarchical folders:

- Sub-folders are directories at the same level as `new/`, `cur/`, `tmp/`
- Sub-folder directory names begin with `.` (dot prefix)
- The hierarchy separator is `.`
- Each sub-folder is a complete Maildir (has its own `new/`, `cur/`, `tmp/`)

```
~/Maildir/
├── new/         <- INBOX new messages
├── cur/         <- INBOX read messages
├── tmp/
├── .Sent/
│   ├── new/
│   ├── cur/
│   └── tmp/
├── .Drafts/
├── .Trash/
└── .Work.Projects.Alpha/   <- INBOX/Work/Projects/Alpha
    ├── new/
    ├── cur/
    └── tmp/
```

**Folder naming rules for Dovecot:**

1. The INBOX maps to the root `~/Maildir/` directory (no `.` prefix)
2. Each sub-folder gets a `.`-prefixed directory
3. Hierarchy nesting uses `.` as separator, e.g., `.Folder.SubFolder`
4. Folder names should not contain:
   - Leading or trailing dots (beyond the structural prefix dot)
   - Slash `/` or backslash `\`
   - Control characters
5. Spaces are technically allowed by most servers but safest avoided; replace with `_`
6. Dots within a folder *name itself* (not as separator) must be escaped or replaced

### 3.4 Dovecot-Specific Files

When Dovecot first accesses a Maildir, it creates its own index files. For an import,
you should **not** pre-create these; Dovecot will regenerate them. What you should
create:

| File | Location | Purpose |
|---|---|---|
| `subscriptions` | `~/Maildir/subscriptions` | List of subscribed folder names (one per line, no leading dot) |
| _(Dovecot creates)_ | `dovecot.index` | Message index |
| _(Dovecot creates)_ | `dovecot-uidlist` | IMAP UID assignment |

The `subscriptions` file should list every folder by its Maildir++ name without the
leading dot. Example:
```
Sent
Trash
Drafts
Work.Projects.Alpha
```

### 3.5 RFC 2822 Message Format

Each file in `cur/` or `new/` must be a valid RFC 2822 message:

```
Date: Mon, 10 Jun 2024 14:30:00 +0000
From: sender@example.com
To: recipient@example.com
Subject: Hello World
Message-ID: <unique-id@example.com>
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="boundary-string"

--boundary-string
Content-Type: text/plain; charset=utf-8

Body text here.

--boundary-string
Content-Type: application/pdf
Content-Disposition: attachment; filename="document.pdf"
Content-Transfer-Encoding: base64

<base64 data>
--boundary-string--
```

---

## 4. Technical Challenges

### 4.1 Challenge Matrix

| Challenge | Severity | Root Cause | Solution |
|---|---|---|---|
| Missing transport headers | High | Exchange-internal messages never hit SMTP | Reconstruct from MAPI properties |
| TNEF / winmail.dat | High | Outlook RTF messages wrap attachments in TNEF | tnefparse decode + reconstruct |
| RTF body without HTML | Medium | Older messages or replies stored as RTF only | striprtf conversion + wrap as text/html |
| Recipient SMTP addresses | Medium | PR_DISPLAY_TO gives display names, not addresses | Iterate recipient table for SMTP_ADDRESS |
| OST-internal message IDs | Medium | No Message-ID for Exchange-internal messages | Synthesize deterministic ID from MAPI properties |
| Timezone normalization | Medium | MAPI FILETIME is UTC, RFC 2822 may have offsets | Convert to UTC epoch; store as GMT |
| Non-ASCII folder names | Medium | Unicode folder names common in non-English locales | UTF-8 encode + sanitize for filesystem |
| Character encoding bodies | Medium | Legacy ANSI OST may have mixed codepages | Detect per-message; fallback to cp1252 |
| High encryption OST | Blocker | libpff cannot decode it | Detect at startup, fail with clear message |
| Named MAPI properties | Low | Custom properties have dynamic IDs | Read Named Property Map at startup |
| Attachment filenames | Low | Short vs long filename properties | Use PR_ATTACH_LONG_FILENAME (0x3707) first |
| Inline images (CID) | Low | HTML body references inline attachments by CID | Preserve Content-ID headers on attachments |
| Duplicate messages | Low | Failed syncs create duplicates | Deduplicate by Message-ID |
| Calendar recurrence | Low | Complex RRULE patterns from MAPI recurrence blob | Parse standard patterns; log and skip complex |
| Calendar exceptions | Low | MAPI stores modified instances as exception blobs | Generate individual VEVENT entries instead |

### 4.2 Reconstructing Missing Internet Headers

When `transport_headers` is `None` or empty, the following MAPI properties provide
the minimum required header set:

```
PR_SUBJECT          (0x0037) → Subject:
PR_SENDER_NAME      (0x0C1A) → From: display name
PR_SENDER_EMAIL_ADDRESS (0x0C1F) / PR_SENDER_SMTP_ADDRESS → From: email
PR_SENT_REPRESENTING_NAME (0x0042) → From: (when sent on behalf of)
PR_CLIENT_SUBMIT_TIME (0x0039) → Date:
PR_MESSAGE_DELIVERY_TIME (0x0E06) → Date: (for received messages)
PR_DISPLAY_TO       (0x0E04) → To: display names
PR_DISPLAY_CC       (0x0E03) → Cc: display names
recipient table     → actual SMTP addresses for each recipient
```

The recipient table in pypff provides individual recipient objects with:
- Display name (`PR_DISPLAY_NAME`)
- SMTP address (`PR_SMTP_ADDRESS` or `PR_EMAIL_ADDRESS`)
- Recipient type (`PR_RECIPIENT_TYPE`: 1=TO, 2=CC, 3=BCC)

### 4.3 TNEF / winmail.dat Handling

Microsoft Outlook encodes messages sent in Rich Text Format (RTF) or using Outlook-
specific features (voting buttons, meeting requests, custom forms) in TNEF format.
The TNEF blob appears as an attachment named `winmail.dat` with MIME type
`application/ms-tnef`.

The TNEF blob contains:
- An RTF or compressed RTF body (may replace or supplement the MIME body)
- Real file attachments that are invisible without TNEF decoding
- MAPI properties that may override the RFC 2822 headers

**Decoding strategy:**
1. Detect `winmail.dat` attachment by name and MIME type
2. Decode with `tnefparse`
3. Extract body (RTF → convert to HTML/text with striprtf)
4. Extract real attachments; add them to the message
5. Remove the `winmail.dat` from the final message
6. If the TNEF body is richer than the existing body, substitute it

### 4.4 Synthetic Message-ID Generation

For messages without a `Message-ID` header, generate a deterministic one to ensure
deduplication works consistently across multiple runs:

```python
import hashlib

def synthetic_message_id(msg, hostname="ost2maildir.local"):
    components = [
        str(msg.identifier),        # internal OST node ID
        str(msg.client_submit_time),
        msg.subject or "",
        msg.sender_name or "",
    ]
    digest = hashlib.sha256("|".join(components).encode()).hexdigest()[:16]
    return f"<{digest}@{hostname}>"
```

This is deterministic: the same message will get the same ID on repeated runs,
so deduplication works even if the output directory already contains partial results.

### 4.5 MAPI FILETIME to RFC 2822 Date

MAPI stores timestamps as Windows FILETIME: a 64-bit integer counting 100-nanosecond
intervals since January 1, 1601, UTC:

```python
import datetime

EPOCH_DIFF = 116444736000000000  # 100ns intervals from 1601 to 1970

def filetime_to_datetime(filetime_int):
    unix_100ns = filetime_int - EPOCH_DIFF
    unix_secs = unix_100ns / 10_000_000
    return datetime.datetime.fromtimestamp(unix_secs, tz=datetime.timezone.utc)

def datetime_to_rfc2822(dt):
    from email.utils import format_datetime
    return format_datetime(dt, usegmt=True)
```

pypff already exposes times as Python `datetime` objects (UTC), so direct FILETIME
arithmetic is only needed when reading raw MAPI properties.

---

## 5. Architecture

### 5.1 Data Flow

```
┌──────────────┐
│   OST File   │
└──────┬───────┘
       │  pypff.file.open()
       ▼
┌──────────────────┐    detect encryption type
│   OSTReader      │─── fail if high encryption
│   (reader.py)    │    iterate folder tree recursively
└──────┬───────────┘
       │  yields (folder_path, item_type, raw_item)
       ▼
┌───────────────────────────────────┐
│   ItemDispatcher (converter.py)   │
│   routes by PR_MESSAGE_CLASS      │
└────┬──────────┬──────────┬────────┘
     │          │          │
     ▼          ▼          ▼
┌─────────┐ ┌─────────┐ ┌───────────┐
│  MIME   │ │  iCal   │ │  vCard    │
│Converter│ │Converter│ │ Converter │
└────┬────┘ └────┬────┘ └────┬──────┘
     │           │            │
     │  if winmail.dat present:
     ▼           │            │
┌─────────────┐  │            │
│TNEFHandler  │  │            │
│(tnef.py)    │  │            │
└────┬────────┘  │            │
     │           │            │
     └─────┬─────┘────────────┘
           │  RFC 2822 email.message.Message objects
           ▼
┌──────────────────────┐
│  Deduplicator        │
│  (dedup.py)          │── Message-ID set (in-memory + persist to .seen file)
└──────────┬───────────┘
           │
           ▼
┌──────────────────────┐
│  MaildirWriter       │── creates folder directories + new/cur/tmp
│  (maildir.py)        │── sanitizes folder names
└──────────┬───────────┘── writes files atomically (tmp → cur)
           │── creates subscriptions file
           ▼
┌──────────────────────┐
│  Output Maildir++    │
│  ~/output/Maildir/   │
└──────────────────────┘
```

### 5.2 Module Structure

```
ost2maildir/
├── ost2maildir/
│   ├── __init__.py
│   ├── cli.py              # Click CLI: entry point, argument parsing, progress
│   ├── reader.py           # OSTReader: pypff wrapper, folder tree iterator
│   ├── converter.py        # MIMEConverter: MAPI → email.message.Message
│   ├── tnef.py             # TNEFHandler: winmail.dat decode via tnefparse
│   ├── rtf.py              # RTFConverter: RTF body → plain text via striprtf
│   ├── calendar.py         # CalendarConverter: IPM.Appointment → .ics
│   ├── contacts.py         # ContactConverter: IPM.Contact → .vcf
│   ├── maildir.py          # MaildirWriter: write messages to Maildir++ tree
│   ├── dedup.py            # Deduplicator: Message-ID deduplication
│   ├── folders.py          # FolderMapper: OST names → Maildir++ names
│   └── utils.py            # Time conversion, encoding helpers, sanitizers
├── tests/
│   ├── fixtures/           # Small synthetic OST/MAPI test data
│   ├── test_converter.py
│   ├── test_tnef.py
│   ├── test_maildir.py
│   ├── test_folders.py
│   └── test_dedup.py
├── pyproject.toml
├── requirements.txt
└── README.md
```

### 5.3 Component Responsibilities

**`reader.py` — OSTReader**

```python
class OSTReader:
    def __init__(self, path: str): ...

    def detect_encryption(self) -> str:
        # Returns "none", "compressible", or "high"
        # Raises EncryptionError if "high"

    def iter_items(self) -> Iterator[tuple[list[str], pypff.item]]:
        # Yields (folder_path_parts, item) for every item in the tree
        # folder_path_parts: ["Inbox"], ["Work", "Projects", "Alpha"], etc.
        # Skips FAI messages (internal Outlook metadata)
        # Handles orphaned nodes (lost+found style recovery)
```

**`converter.py` — MIMEConverter**

```python
class MIMEConverter:
    def convert(self, msg: pypff.message) -> email.message.Message:
        # 1. Try transport_headers as base; supplement with MAPI properties
        # 2. Build body: HTML > plain text > RTF-converted
        # 3. Attach file attachments
        # 4. If winmail.dat detected, hand off to TNEFHandler
        # 5. Ensure required headers: Date, From, To, Subject, Message-ID
        # 6. Normalize all header encodings to RFC 2047 UTF-8

    def _get_recipients(self, msg) -> dict[str, list[str]]:
        # Returns {"to": [...], "cc": [...], "bcc": [...]}
        # Builds "Display Name <smtp@addr>" strings
```

**`maildir.py` — MaildirWriter**

```python
class MaildirWriter:
    def __init__(self, output_dir: str): ...

    def write(self, folder_path: list[str], mime_msg: email.message.Message,
              flags: str = "") -> None:
        # Ensures directory exists (new/cur/tmp)
        # Generates unique filename with timestamp + identifier
        # Writes to tmp/ first, then renames to cur/ (atomic)

    def finalize(self) -> None:
        # Writes subscriptions file
        # Logs summary statistics
```

---

## 6. Dependency Stack

### 6.1 Required Dependencies

| Package | PyPI name | Version | Purpose | Risk |
|---|---|---|---|---|
| **pypff** | `pypff` | from libpff source | OST/PST reading | Must build from source (no stable PyPI wheel) |
| **tnefparse** | `tnefparse` | ≥ 1.4.0 | TNEF/winmail.dat decode | Marked inactive but stable, no known vulnerabilities |
| **striprtf** | `striprtf` | ≥ 0.0.26 | RTF body → plain text | Active, maintained |
| **icalendar** | `icalendar` | ≥ 5.0 | iCal (.ics) generation | Active, well-maintained |
| **vobject** | `vobject` | ≥ 0.9.6 | vCard (.vcf) generation | Slow development but stable API |
| **click** | `click` | ≥ 8.0 | CLI argument parsing | Production-grade, part of Flask ecosystem |
| **tqdm** | `tqdm` | ≥ 4.0 | Progress bars | Ubiquitous, stable |
| **chardet** | `chardet` | ≥ 5.0 | Fallback charset detection | Stable |

### 6.2 Installing pypff (libpff Python Bindings)

**pypff is not on PyPI as a maintained wheel.** It must be built from the libpff source:

```bash
# Prerequisites (Debian/Ubuntu/WSL2)
sudo apt install build-essential python3-dev libfuse-dev

git clone https://github.com/libyal/libpff.git
cd libpff
./synclibs.sh
./autogen.sh
./configure --enable-python3
make
sudo make install
sudo ldconfig

# Build and install the Python binding
cd pypff
python3 setup.py build
sudo python3 setup.py install
```

Alternative (often works):
```bash
pip install libpff-python
```
(This is a community-maintained wheel; verify the source before using.)

### 6.3 Alternative: libratom

`libratom` (https://pypi.org/project/libratom/) wraps pypff with a higher-level API
and can export to `.eml` format. Consider using it as a foundation layer for message
reading, then post-process the `.eml` output into Maildir++. This trades flexibility
for reliability on the parsing layer.

### 6.4 Full requirements.txt

```
# Core
libpff-python>=20221124     # or install from source; see §6.2
tnefparse>=1.4.0
striprtf>=0.0.26
icalendar>=5.0.13
vobject>=0.9.6.1

# CLI and UX
click>=8.1.0
tqdm>=4.66.0

# Encoding
chardet>=5.2.0

# Dev/test only
pytest>=8.0
pytest-cov>=5.0
```

---

## 7. Property Mappings

### 7.1 MAPI Property → RFC 2822 Header Mapping

| MAPI Property ID | Property Name | RFC 2822 Header | Notes |
|---|---|---|---|
| `0x0037` | `PR_SUBJECT` | `Subject:` | Unicode or ANSI |
| `0x0C1A` | `PR_SENDER_NAME` | `From:` display name | |
| `0x0C1F` | `PR_SENDER_EMAIL_ADDRESS` | `From:` address | May be X.500 DN; prefer SMTP |
| `0x5D02` | `PR_SENDER_SMTP_ADDRESS` | `From:` address | Prefer over 0x0C1F |
| `0x0042` | `PR_SENT_REPRESENTING_NAME` | `From:` | When sent on behalf |
| `0x0039` | `PR_CLIENT_SUBMIT_TIME` | `Date:` | FILETIME → RFC 2822 |
| `0x0E06` | `PR_MESSAGE_DELIVERY_TIME` | `Date:` (received) | Use for received mail |
| `0x1035` | `PR_INTERNET_MESSAGE_ID` | `Message-ID:` | If absent, synthesize |
| `0x1042` | `PR_IN_REPLY_TO_ID` | `In-Reply-To:` | Threading |
| `0x1039` | `PR_INTERNET_REFERENCES` | `References:` | Threading chain |
| `0x0E04` | `PR_DISPLAY_TO` | `To:` (fallback) | Display names only; use recipient table for addresses |
| `0x0E03` | `PR_DISPLAY_CC` | `Cc:` (fallback) | Display names only |
| `0x0017` | `PR_IMPORTANCE` | `X-Priority:` / `Importance:` | 0=Low, 1=Normal, 2=High |
| `0x0036` | `PR_SENSITIVITY` | `Sensitivity:` | 0=Normal, 1=Personal, 2=Private, 3=Confidential |
| `0x0070` | `PR_CONVERSATION_TOPIC` | `X-MS-Conversation-Topic:` | Thread grouping |
| `0x007D` | `PR_TRANSPORT_MESSAGE_HEADERS` | (all headers) | Use as header block base if present |
| `0x1000` | `PR_BODY` | `text/plain` body part | |
| `0x1013` | `PR_HTML` | `text/html` body part | Raw bytes; detect charset from HTML `<meta>` |
| `0x1009` | `PR_RTF_COMPRESSED` | (convert → text/html) | Decompressed by libpff; convert with striprtf |
| `0x3707` | `PR_ATTACH_LONG_FILENAME` | attachment `filename=` | Prefer over short name |
| `0x3704` | `PR_ATTACH_FILENAME` | attachment `filename=` | 8.3 fallback only |
| `0x3703` | `PR_ATTACH_EXTENSION` | attachment MIME type | Derive MIME type from extension if needed |
| `0x3712` | `PR_ATTACH_CONTENT_ID` | `Content-ID:` on attachment | For inline image references |

**Recipient Table Properties (per recipient entry):**

| Property | Meaning |
|---|---|
| `PR_DISPLAY_NAME (0x3001)` | Display name |
| `PR_SMTP_ADDRESS (0x39FE)` | Internet email address (preferred) |
| `PR_EMAIL_ADDRESS (0x3003)` | May be X.500 format; skip if X.500 |
| `PR_RECIPIENT_TYPE (0x0C15)` | 1=TO, 2=CC, 3=BCC |
| `PR_ADDRTYPE (0x3002)` | "SMTP", "EX" (Exchange X.500), "MAPIPDL" (dist list) |

Only use the address if `PR_ADDRTYPE` is "SMTP". Otherwise use display name only.

### 7.2 MAPI Message Flags → Maildir Flags

`PR_MESSAGE_FLAGS` (0x0E07) is a bitmask:

| Bitmask | Constant | Maildir Flag | Notes |
|---|---|---|---|
| `0x0001` | `MSGFLAG_READ` | `S` (Seen) | Inverse: if NOT set → message goes to `new/` |
| `0x0010` | `MSGFLAG_HASATTACH` | (none) | No Maildir equivalent |
| `0x0020` | `MSGFLAG_FROMME` | (none) | Indicates sender is the mailbox owner |
| `0x0100` | `MSGFLAG_UNSENT` | `D` (Draft) | Message never sent |
| `0x0200` | `MSGFLAG_RESEND` | (none) | |

**Follow-up/Flagged:** Outlook's "Follow Up" flag is stored as a combination of:
- Named property `PR_FLAG_STATUS` (GUID: {00062003-...}, ID: 0x1090): 1=flagged, 2=complete
- If flagged → add `F` to Maildir flags

**Replied/Forwarded:** Outlook tracks these in:
- `PR_LAST_VERB_EXECUTED` (0x1081): 102=reply, 103=reply-all, 104=forward
- 102 or 103 → add `R` (Replied) to Maildir flags
- 104 → add `P` (Passed/Forwarded) to Maildir flags

**Placement rule:** Messages with `S` (Seen) flag go to `cur/`. Messages without any
flags (unread) go to `new/`. For import, place all messages in `cur/` — the IMAP
server will expose the Seen flag state to the client through the filename.

### 7.3 Outlook Folder Names → Maildir++ Names

**Standard folder mapping (case-insensitive match on OST folder name):**

| Outlook Folder Name | Maildir++ Path | Notes |
|---|---|---|
| `Inbox` | (root `~/Maildir/`) | INBOX is the Maildir root |
| `Sent Items` | `.Sent` | Dovecot default Sent folder |
| `Deleted Items` | `.Trash` | Dovecot default Trash |
| `Drafts` | `.Drafts` | Dovecot default Drafts |
| `Junk E-mail` | `.Junk` | Also match `Spam`, `Junk Mail` |
| `Outbox` | `.Outbox` | |
| `Archive` | `.Archive` | |
| `Sent` | `.Sent` | Some tenants use this name |
| `Calendar` | `.Calendar` | Used for converted appointments |
| `Contacts` | `.Contacts` | Used for converted contacts |
| (any other) | `.{sanitized_name}` | Custom folder |

**Nested folder example:**
OST: `Inbox > Work > Projects > Alpha`
Maildir++: `.Work.Projects.Alpha` (Inbox content stays in root)

**Folder name sanitization:**
```python
import re

def sanitize_folder_name(name: str) -> str:
    # Remove or replace characters unsafe for Maildir++ and filesystems
    name = name.strip()
    name = re.sub(r'[/\\:*?"<>|]', '_', name)  # filesystem-unsafe
    name = name.replace('.', '_')               # dots are hierarchy separators
    name = re.sub(r'\s+', '_', name)            # spaces → underscores
    name = name.strip('_')
    return name or "unnamed"
```

---

## 8. Implementation Plan (Phased)

### Phase 1 — Foundation & Environment

**Goal:** Prove the OST can be opened and iterated.

1. Set up Python project with `pyproject.toml`, virtual environment
2. Install and verify libpff/pypff builds correctly on target system
3. Implement `OSTReader`:
   - Open file with encryption detection
   - Recursive folder tree walker yielding `(path_parts, item)`
   - Count folders and items, print summary (dry-run mode)
4. Print a folder tree to stdout for verification:
   ```
   Top of Information Store
   ├── Inbox (3,241 messages)
   ├── Sent Items (1,102 messages)
   ├── Deleted Items (445 messages)
   └── Work/
       └── Projects/ (88 messages)
   ```
5. Identify and skip FAI (folder-associated information) items

**Verification:** Run against a real OST file and produce an accurate folder/count summary.

---

### Phase 2 — Basic MIME Conversion

**Goal:** Convert a single email message to a valid RFC 2822 `.eml` file.

1. Implement `MIMEConverter.convert()`:
   - Route: if `transport_headers` present → parse as header block
   - Supplement with MAPI properties for any missing required headers
   - Build body: prefer HTML, fall back to plain text, fall back to RTF-converted
   - Add `MIME-Version: 1.0` and `Content-Type`
2. Implement `RTFConverter` using `striprtf`:
   - Wrap result in `text/plain; charset=utf-8`
3. Implement recipient extraction from MAPI recipient table
4. Implement synthetic Message-ID generation
5. Implement `utils.py`:
   - FILETIME → datetime (if needed)
   - RFC 2047 header encoding for non-ASCII values
   - Charset detection for bytes bodies
6. Write 5–10 messages to individual `.eml` files for manual inspection

**Verification:** Open converted `.eml` files in Thunderbird or `mutt`; verify headers,
body, dates are correct.

---

### Phase 3 — Attachment Handling

**Goal:** Messages with attachments convert correctly.

1. Extend `MIMEConverter` to build `multipart/mixed` messages
2. Stream attachment data to avoid OOM: use `att.read_buffer()` in chunks
3. Set `Content-Disposition: attachment; filename=...` using `PR_ATTACH_LONG_FILENAME`
4. Preserve `Content-ID` on inline attachments (for HTML `<img src="cid:...">`)
5. Handle embedded message attachments (`.msg` files, IPM.Note nested in attachment)
6. Detect MIME type from file extension using Python `mimetypes` module

**Verification:** Convert messages with known attachment types (PDF, DOCX, PNG); verify
attachments are present and correct in resulting `.eml`.

---

### Phase 4 — TNEF / winmail.dat Decoding

**Goal:** Exchange-sent messages with hidden attachments are decoded.

1. Implement `TNEFHandler`:
   - Detect: attachment named `winmail.dat` with type `application/ms-tnef`
   - Decode with `tnefparse.TNEF(data)`
   - Extract `.attachments` → re-attach each to the message
   - Extract `.body` (RTF) or `.htmlbody` → substitute for message body if richer
   - Remove the original `winmail.dat` attachment from the message
2. Handle edge case: TNEF with no attachments (body-only TNEF); still decode body
3. Log TNEF decode failures as warnings; preserve winmail.dat on failure

**Verification:** Source a real winmail.dat-containing message; verify hidden attachments
appear in output `.eml`.

---

### Phase 5 — Maildir++ Output

**Goal:** All converted messages land in the correct Maildir++ folder structure.

1. Implement `FolderMapper`:
   - Standard folder name mapping table (§7.3)
   - Recursive path sanitization
   - Generate Dovecot-compatible `.`-separated path
2. Implement `MaildirWriter`:
   - Create `new/`, `cur/`, `tmp/` on first write to a folder
   - Generate unique filenames: `{timestamp}.{identifier}.{hostname}:2,{flags}`
   - Derive flags from MAPI `PR_MESSAGE_FLAGS` and follow-up properties (§7.2)
   - Write to `tmp/` then `os.rename()` to `cur/` (atomic, POSIX-safe)
   - Track all folder names for `subscriptions` file
3. Implement `Deduplicator`:
   - In-memory set of seen Message-IDs per run
   - Persist seen set to `{output}/.ost2maildir_seen` for resumable runs
   - Log duplicates skipped
4. Write `subscriptions` file after all messages are processed
5. Implement progress bar via `tqdm` across all messages

**Verification:** Import output into a live Dovecot instance with `doveadm import`;
verify folder structure and message counts match expected.

---

### Phase 6 — Calendar and Contacts

**Goal:** Appointments and contacts appear as importable items in the Maildir.

1. Implement `CalendarConverter`:
   - Detect `IPM.Appointment` and `IPM.Schedule.Meeting.*` classes
   - Map MAPI properties to VEVENT fields (§8.1 below)
   - Output as `text/calendar` attachment named `event.ics` in a
     `multipart/mixed` message with subject `[Calendar] {title}`
   - Place in `.Calendar` folder
2. Implement `ContactConverter`:
   - Detect `IPM.Contact` class
   - Map MAPI properties to vCard 3.0 fields (§8.2 below)
   - Output as `text/vcard` attachment named `contact.vcf` in a
     `multipart/mixed` message with subject `[Contact] {full_name}`
   - Place in `.Contacts` folder
3. Handle `IPM.DistList` (distribution lists) → vCard GROUP type

**Calendar MAPI → iCal mapping (§8.1):**

| MAPI Property | iCal Field |
|---|---|
| `PR_SUBJECT (0x0037)` | `SUMMARY` |
| `PR_BODY (0x1000)` | `DESCRIPTION` |
| `PR_START_DATE (0x0060)` | `DTSTART` |
| `PR_END_DATE (0x0061)` | `DTEND` |
| `PR_LOCATION (0x0C1E Appt)` | `LOCATION` |
| `PR_CREATION_TIME` | `CREATED` |
| `PR_LAST_MODIFICATION_TIME` | `LAST-MODIFIED` |
| `PR_SENSITIVITY` | `CLASS` (PUBLIC/PRIVATE/CONFIDENTIAL) |
| `PR_IMPORTANCE` | `PRIORITY` |
| `PR_RECURRENCE_PATTERN (0x3006)` | `RRULE` (simple patterns only) |
| `PR_ORGANIZER (named prop)` | `ORGANIZER` |

**Contact MAPI → vCard mapping (§8.2):**

| MAPI Property | vCard Field |
|---|---|
| `PR_GIVEN_NAME (0x3A06)` + `PR_SURNAME (0x3A11)` | `N:` / `FN:` |
| `PR_COMPANY_NAME (0x3A16)` | `ORG:` |
| `PR_TITLE (0x3A17)` | `TITLE:` |
| `PR_BUSINESS_TELEPHONE_NUMBER (0x3A08)` | `TEL;TYPE=WORK:` |
| `PR_HOME_TELEPHONE_NUMBER (0x3A09)` | `TEL;TYPE=HOME:` |
| `PR_MOBILE_TELEPHONE_NUMBER (0x3A1C)` | `TEL;TYPE=CELL:` |
| `PR_EMAIL_ADDRESS_1 (named)` | `EMAIL;TYPE=INTERNET:` |
| `PR_BUSINESS_ADDRESS (named)` | `ADR;TYPE=WORK:` |
| `PR_HOME_ADDRESS (named)` | `ADR;TYPE=HOME:` |
| `PR_BODY (0x1000)` | `NOTE:` |
| `PR_THUMBNAIL_PHOTO (named)` | `PHOTO;ENCODING=b:` |

---

### Phase 7 — CLI, Logging, and Polish

**Goal:** Production-ready tool with clear UX.

1. Implement full CLI with `click` (§9)
2. Add structured logging: INFO to stderr, DEBUG to file
3. Add summary report at completion:
   - Messages converted / skipped (duplicate) / failed
   - Attachments extracted
   - TNEF blobs decoded
   - Calendar items / contacts written
   - Folders created
4. Add `--dry-run` mode: parse without writing
5. Add `--folder-filter` glob pattern support
6. Package with `pyproject.toml` entry point

---

## 9. CLI Design and Usage

### 9.1 Command Signature

```
ost2maildir [OPTIONS] INPUT_OST OUTPUT_MAILDIR
```

### 9.2 Options

```
Arguments:
  INPUT_OST       Path to the .ost file to convert.
  OUTPUT_MAILDIR  Output directory (created if absent). This becomes ~/Maildir/
                  on the target server.

Options:
  --no-dedup             Disable Message-ID deduplication.
  --skip-contacts        Do not convert contacts to .vcf messages.
  --skip-calendar        Do not convert calendar items to .ics messages.
  --no-tnef              Skip TNEF decoding (preserve winmail.dat as-is).
  --folder-filter TEXT   Only process folders whose full path matches this
                         glob pattern (e.g. "Inbox/*", "Work/**").
  --encoding TEXT        Default character encoding for ANSI bodies
                         [default: cp1252].
  --log-file PATH        Write full conversion log to this file.
  --dry-run              Parse and count items without writing output.
  --hostname TEXT        Hostname to use in generated Message-IDs and
                         Maildir filenames [default: system hostname].
  -v, --verbose          Enable debug-level logging.
  --version              Show version and exit.
  --help                 Show this message and exit.
```

### 9.3 Usage Examples

**Basic full-mailbox conversion:**
```bash
ost2maildir outlook.ost ~/Maildir
```

**Dry run to assess the file:**
```bash
ost2maildir --dry-run outlook.ost /tmp/unused
```
Output:
```
OST encryption: compressible (readable)
Folders found: 24
Messages found: 14,382
Calendar items: 891
Contacts: 234
Estimated output: ~4.2 GB
```

**Convert only the Inbox and Sent Items:**
```bash
ost2maildir --folder-filter "Inbox/**" --folder-filter "Sent Items/**" \
    outlook.ost ~/Maildir
```

**Convert with full logging:**
```bash
ost2maildir --log-file conversion.log -v outlook.ost ~/Maildir
```

**Skip calendar and contacts, preserve winmail.dat:**
```bash
ost2maildir --skip-calendar --skip-contacts --no-tnef outlook.ost ~/Maildir
```

### 9.4 Importing into Dovecot

After conversion, import using Dovecot's `doveadm`:

```bash
# Option 1: Copy the Maildir directly into the user's mail location
rsync -av ~/Maildir/ /var/mail/vhosts/example.com/username/

# Option 2: Use doveadm import (preserves UIDs, handles duplicates)
doveadm import -u username@example.com maildir:/home/username/Maildir "" "*"

# Option 3: dsync (for replication or full migration)
doveadm sync -u username@example.com maildir:/home/username/Maildir
```

After import, tell Dovecot to rebuild its indexes:
```bash
doveadm index -u username@example.com INBOX
doveadm index -u username@example.com "*"
```

---

## 10. Configuration Reference

No configuration file is needed for a single-run tool. All behavior is controlled via
CLI flags (§9.2). The only persistent state the tool writes is:

| File | Location | Purpose |
|---|---|---|
| `.ost2maildir_seen` | `OUTPUT_MAILDIR/` | Message-ID dedup set (JSON); enables resumable runs |
| `subscriptions` | `OUTPUT_MAILDIR/` | Dovecot folder subscription list |
| `conversion.log` | user-specified | Full conversion log (if `--log-file` used) |

---

## 11. Caveats and Known Limitations

### 11.1 High Encryption OST Files (Critical)

OST files with high encryption (`0x02`) **cannot be read** by libpff or any open-source
tool. High encryption uses a proprietary multi-rotor substitution cipher. If your OST
is high-encrypted, you must:
1. Open the OST in Outlook (requires the original Exchange/M365 account)
2. Export to PST via File > Open & Export > Import/Export > "Export to a file"
3. Run `ost2maildir` on the resulting PST (identical file format)

The tool will detect high encryption at startup and print a clear error with
instructions rather than producing garbled output.

### 11.2 Compressible Encryption (Normal Operation)

The default OST encryption mode (`0x01`) is trivially reversed by libpff
automatically. No user action is required.

### 11.3 Reconstructed Headers May Differ from Originals

For Exchange-internal messages (no `transport_headers`), headers are reconstructed
from MAPI properties. The reconstructed message is semantically complete but may
differ from what an SMTP relay would have generated:
- `Received:` chain headers are absent
- `Return-Path:` may be absent
- `X-` vendor headers are absent
- `Message-ID` may be synthesized (deterministically, based on OST node ID + timestamp)

For forensic or legal use cases where bit-perfect header fidelity is required, this
tool is not suitable. For normal email migration, the reconstruction is complete.

### 11.4 RTF Body Conversion Quality

When only `PR_RTF_COMPRESSED` is available (no HTML or plain-text body), `striprtf`
converts RTF to plain text. Complex RTF documents — tables, embedded objects, complex
formatting — will lose formatting. The text content is preserved. If the OST also
contains `PR_HTML`, that is always preferred and this limitation does not apply.

### 11.5 Calendar Recurrence Patterns

MAPI stores recurrence as a binary blob (`PR_RECURRENCE_PATTERN`, property
`0x80E1`). libpff does not decode this blob; the tool must parse it manually or use
a helper library. The following patterns are straightforward to convert to iCal
RRULE:
- Daily, every N days
- Weekly, on specific days
- Monthly, on a specific day number
- Yearly, on a specific month+day

The following are complex and will be approximated or expanded to individual events:
- "First/last weekday of month"
- "Every Nth weekday of month"
- Exception-modified instances (e.g., "third Tuesday, but moved to Wednesday on
  Feb 14")
- Recurrence patterns with end-date exceptions

A warning is logged for any pattern that cannot be expressed as a clean iCal RRULE;
the event is still exported as a non-recurring VEVENT.

### 11.6 Maximum Attachment Size and Memory

pypff's `att.data` property loads the entire attachment into memory. For attachments
over ~50 MB, use the streaming `att.read_buffer(chunk_size)` API instead. The tool
should always use streaming reads for attachments, writing to `tmp/` incrementally
before linking to `cur/`.

OST files from very heavy users may contain individual attachments of 1 GB or more
(e.g., Outlook PST files attached to messages). The tool handles these via streaming
but output disk space must be at least as large as the total attachment volume.

### 11.7 Contact Address Types

Outlook contacts may store email addresses using Exchange X.500 Distinguished Names
(`/o=ExchangeOrg/ou=Exchange Administrative Group/...`) rather than SMTP addresses.
X.500 addresses cannot be imported to a standard mail client. The tool will:
- Skip non-SMTP address types in the `From:` and recipient headers
- Log a warning for each X.500-only contact
- Preserve the display name even when the address is absent

### 11.8 Unicode / Multibyte Folder Names

MAPI folder names in modern OST files are UTF-16LE. pypff decodes these to Python
`str` (Unicode). Filesystem constraints on the target may cause issues if:
- The output is on a Windows filesystem (NTFS via WSL2): certain characters are
  forbidden (`:`, `*`, `?`, `"`, `<`, `>`, `|`)
- The output is on a Samba share with specific encoding settings

The `FolderMapper` sanitizes names before use (§7.3) and the original name is
preserved in the folder structure's UIDL file for reference.

### 11.9 Orphaned and Recovered Items

OST files may contain "orphaned" nodes — messages or folders not attached to the
main folder tree (e.g., from interrupted sync operations). pffexport calls these
`*.orphans`. The tool can optionally recover these into a `.Recovered` folder at the
output root. This is disabled by default; enable with `--recover-orphans`.

### 11.10 Large Files and `ulimit`

Converting a 20 GB OST with 50,000 messages will create 50,000+ small files.
Default Linux `ulimit -n` is 1024 open file descriptors. The tool opens and closes
files one at a time (no parallel writes), so this is not a hard limit during
conversion. However, `inotify` watches during Dovecot indexing may hit limits on
large Maildir imports. On the Dovecot server:

```bash
echo "fs.inotify.max_user_watches = 524288" >> /etc/sysctl.conf
sysctl -p
```

### 11.11 Exchange Distribution Lists (DL)

`IPM.DistList` objects represent Outlook distribution lists. These are converted to
vCard GROUP type. However, the member list may contain X.500 addresses (see §11.7)
that will be filtered out, potentially resulting in incomplete vCard groups.

### 11.12 S/MIME and PGP Encrypted Messages

If the mailbox contains messages encrypted with S/MIME or PGP, the ciphertext is
preserved in the output — the tool does not attempt decryption. Encrypted attachments
will appear as opaque `application/pkcs7-mime` or `application/pgp-encrypted` parts.

---

## 12. Testing Strategy

### 12.1 Unit Tests

Test each converter component in isolation with synthetic MAPI-like inputs:

```python
# test_converter.py
def test_subject_with_unicode():
    """Korean characters in subject survive round-trip through RFC 2047 encoding."""

def test_missing_transport_headers_reconstructed():
    """Message without transport_headers gets valid From/To/Date headers."""

def test_synthetic_message_id_is_deterministic():
    """Same input produces same Message-ID on repeated calls."""

def test_mapi_flags_to_maildir_flags():
    """MSGFLAG_READ + follow-up flag → 'FS' Maildir flags."""
```

### 12.2 Integration Tests

Use real (anonymized) OST fragments or the `pfftools` test data available in the
libpff repository (`libpff/tests/input/`).

Key integration test scenarios:
- OST with compressible encryption → opens successfully
- OST with high encryption → raises `EncryptionError` with clear message
- Message with only RTF body → `striprtf` produces readable plain text
- Message with `winmail.dat` → TNEF decoded, real attachments visible
- Deeply nested folder hierarchy (5 levels) → correct Maildir++ paths
- 1000-message folder → all messages written, no file naming collisions
- Duplicate messages → only one written when deduplication enabled
- Calendar item with daily recurrence → valid iCal RRULE in output
- Contact with JPEG photo → `PHOTO` field present in vCard output

### 12.3 Validation Tools

After conversion, validate output with:

```bash
# Validate all messages are valid RFC 2822
for f in $(find ~/Maildir -name '*:2,*'); do
    python3 -c "import email; email.message_from_file(open('$f'))" || echo "INVALID: $f"
done

# Validate iCal files
python3 -c "
import icalendar, pathlib
for f in pathlib.Path('Maildir/.Calendar/cur').glob('*:2,*'):
    raw = f.read_bytes()
    import email; msg = email.message_from_bytes(raw)
    for part in msg.walk():
        if part.get_content_type() == 'text/calendar':
            icalendar.Calendar.from_ical(part.get_payload(decode=True))
            print(f'OK: {f.name}')
"

# Count messages per folder vs expected from dry-run
find ~/Maildir -name '*:2,*' | wc -l
```

### 12.4 Manual Dovecot Import Test

```bash
# Minimal Dovecot config for local testing
# /etc/dovecot/dovecot.conf:
#   mail_location = maildir:~/Maildir

doveadm import -u testuser maildir:/tmp/converted_maildir "" "*"
doveadm mailbox list -u testuser
doveadm stat -u testuser INBOX
```

---

## 13. Performance Considerations

### 13.1 Expected Throughput

On modern hardware (SSD, 8-core CPU), with a 10 GB OST containing ~30,000 messages:
- OST opening and header scan: 5–15 seconds (libpff B-tree traversal)
- Message conversion: ~50–200 messages/second depending on attachment size
- TNEF decode: adds ~10 ms per message with winmail.dat
- Estimated total: 3–10 minutes for 30,000 messages (excluding large attachments)

Large attachment messages (100+ MB) will dominate total time; each must be streamed.

### 13.2 Memory Profile

The tool should maintain a low memory footprint:
- pypff caches the B-tree index in memory (~10–50 MB for a 10 GB file)
- Message bodies and attachments should be processed and discarded one at a time
- The Message-ID dedup set grows ~100 bytes per message: 1M messages = ~100 MB
- For mailboxes over 500K messages, consider persisting the dedup set to SQLite

### 13.3 Parallelism

The OST reading is single-threaded (libpff is not thread-safe). However, after a
message is read and converted to a Python `email.message.Message` object, writing
to disk can be done in a thread pool:

```python
from concurrent.futures import ThreadPoolExecutor

with ThreadPoolExecutor(max_workers=4) as executor:
    for path, msg, flags in converter.iter_converted():
        executor.submit(writer.write, path, msg, flags)
```

Caution: thread-safety must be enforced in `MaildirWriter` (separate locks per
folder, or pre-create all directories before entering the thread pool).

For v1, single-threaded is recommended to simplify debugging.

### 13.4 Disk Space

Plan for output size ≈ 1.0–1.3× the OST file size. Reasons:
- TNEF decode expands compact TNEF bundles into separate attachment files
- Base64 encoding of binary attachments in MIME adds ~33% overhead
- Maildir per-file overhead for small messages (filesystem block size)

Run `--dry-run` first; the tool will estimate total output size.

---

*End of OST2MAILDIR Technical Design Document*

*Generated: 2026-06-02 | Target: Dovecot Maildir++ on Linux | Language: Python 3.10+*
