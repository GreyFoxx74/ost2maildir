"""Shared helpers: MAPI property access, subject cleaning, charset detection."""

import datetime
import hashlib
import re
import socket
from typing import Any


FILETIME_EPOCH_DIFF = 116_444_736_000_000_000  # 100-ns ticks from 1601 to 1970-01-01


def filetime_to_datetime(filetime: int) -> datetime.datetime:
    unix_100ns = filetime - FILETIME_EPOCH_DIFF
    unix_secs = unix_100ns / 10_000_000
    return datetime.datetime.fromtimestamp(unix_secs, tz=datetime.timezone.utc)


def clean_subject(s: str | None) -> str:
    """Strip MAPI normalized-subject prefix bytes (0x01 0xNN ...) from PR_SUBJECT."""
    if not s:
        return ""
    # MAPI encodes prefix indicator as: byte[0]=0x01, byte[1]=prefix_length, then full subject
    if len(s) >= 2 and s[0] == "\x01":
        s = s[2:]
    # Strip any residual leading control characters
    s = s.lstrip("".join(chr(c) for c in range(0x20)))
    return s.strip()


def is_smtp_address(addr: str) -> bool:
    """Return True if the address looks like an SMTP address, not X.500 / DN."""
    if not addr:
        return False
    addr_upper = addr.upper()
    if addr_upper.startswith("/O="):
        return False
    if addr_upper.startswith("CN="):
        return False
    return "@" in addr


def get_mapi_prop(record_set, prop_id: int) -> Any | None:
    """Return the first value matching prop_id from a pypff record_set, or None."""
    try:
        for i in range(record_set.number_of_entries):
            entry = record_set.get_entry(i)
            if entry.entry_type == prop_id:
                return entry
    except Exception:
        pass
    return None


def get_all_mapi_props(msg) -> dict[int, Any]:
    """Return {prop_id: entry} for all properties in the first record_set."""
    props: dict[int, Any] = {}
    try:
        if msg.number_of_record_sets < 1:
            return props
        rs = msg.get_record_set(0)
        for i in range(rs.number_of_entries):
            entry = rs.get_entry(i)
            props[entry.entry_type] = entry
    except Exception:
        pass
    return props


def entry_as_string(entry) -> str | None:
    """Safely extract a string value from a MAPI record entry."""
    if entry is None:
        return None
    try:
        return entry.data_as_string
    except Exception:
        return None


def entry_as_int(entry) -> int | None:
    if entry is None:
        return None
    try:
        return entry.data_as_integer
    except Exception:
        return None


def entry_as_datetime(entry) -> datetime.datetime | None:
    if entry is None:
        return None
    try:
        dt = entry.data_as_datetime
        if dt and dt.tzinfo is None:
            dt = dt.replace(tzinfo=datetime.timezone.utc)
        return dt
    except Exception:
        return None


def detect_html_charset(html_bytes: bytes) -> str:
    """Extract charset from HTML <meta> tags; fall back to utf-8."""
    try:
        snippet = html_bytes[:4096].decode("ascii", errors="ignore").lower()
        # <meta charset="utf-8">
        m = re.search(r'<meta[^>]+charset=["\']?([a-z0-9_\-]+)', snippet)
        if m:
            return m.group(1)
        # <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
        m = re.search(r'charset=([a-z0-9_\-]+)', snippet)
        if m:
            return m.group(1)
    except Exception:
        pass
    return "utf-8"


def decode_body(body: bytes | str | None, charset: str = "utf-8") -> str | None:
    """Decode a message body to str; return None if body is empty."""
    if body is None:
        return None
    if isinstance(body, str):
        return body if body.strip() else None
    if not body:
        return None
    for enc in (charset, "utf-8", "windows-1252", "latin-1"):
        try:
            return body.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return body.decode("latin-1", errors="replace")


def synthetic_message_id(msg, hostname: str | None = None) -> str:
    """Generate a deterministic Message-ID for messages that lack one."""
    if hostname is None:
        hostname = socket.gethostname() or "ost2maildir.local"
    parts = [
        str(msg.identifier),
        str(msg.client_submit_time or msg.delivery_time or msg.creation_time or ""),
        clean_subject(msg.subject) or "",
        msg.sender_name or "",
    ]
    digest = hashlib.sha256("|".join(parts).encode("utf-8", errors="replace")).hexdigest()[:24]
    return f"<{digest}@{hostname}>"


def mapi_flags_to_maildir(flags_int: int, props: dict) -> str:
    """Map PR_MESSAGE_FLAGS + follow-up props to a sorted Maildir flags string."""
    flags: set[str] = set()

    if flags_int & 0x0001:  # MSGFLAG_READ
        flags.add("S")
    if flags_int & 0x0100:  # MSGFLAG_UNSENT (draft)
        flags.add("D")

    # PR_LAST_VERB_EXECUTED: 102=reply, 103=reply-all, 104=forward
    verb_entry = props.get(0x1081)
    verb = entry_as_int(verb_entry)
    if verb in (102, 103):
        flags.add("R")
    elif verb == 104:
        flags.add("P")

    # Follow-up / flagged: named property PR_FLAG_STATUS is dynamic; approximate via
    # PR_TODO_ITEM_FLAGS (0x0FFC) or PR_FLAG_STATUS pattern
    flag_entry = props.get(0x1090)  # common named-prop ID for flag status
    if flag_entry:
        fval = entry_as_int(flag_entry)
        if fval == 1:  # flagged
            flags.add("F")

    return "".join(sorted(flags))
