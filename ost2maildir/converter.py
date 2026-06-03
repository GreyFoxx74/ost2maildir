"""MAPI → RFC 2822 MIME message converter."""

import datetime
import logging
import mimetypes
import socket
from email import policy
from email.headerregistry import Address
from email.message import EmailMessage
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import format_datetime, formataddr, parseaddr
from typing import Optional

from .tnef import process_tnef
from .utils import (
    clean_subject,
    decode_body,
    detect_html_charset,
    entry_as_int,
    entry_as_string,
    get_all_mapi_props,
    is_smtp_address,
    mapi_flags_to_maildir,
    synthetic_message_id,
)

log = logging.getLogger(__name__)

# MAPI property IDs we care about
PR_SUBJECT = 0x0037
PR_SENDER_NAME = 0x0C1A
PR_SENDER_EMAIL_ADDRESS = 0x0C1F
PR_SENDER_SMTP_ADDRESS = 0x5D02
PR_SENT_REPRESENTING_NAME = 0x0042
PR_SENT_REPRESENTING_EMAIL = 0x0076
PR_CLIENT_SUBMIT_TIME = 0x0039
PR_MESSAGE_DELIVERY_TIME = 0x0E06
PR_INTERNET_MESSAGE_ID = 0x1035
PR_IN_REPLY_TO_ID = 0x1042
PR_INTERNET_REFERENCES = 0x1039
PR_DISPLAY_TO = 0x0E04
PR_DISPLAY_CC = 0x0E03
PR_IMPORTANCE = 0x0017
PR_SENSITIVITY = 0x0036
PR_MESSAGE_FLAGS = 0x0E07
PR_MESSAGE_CLASS = 0x001A
PR_CONVERSATION_TOPIC = 0x0070
PR_LAST_VERB_EXECUTED = 0x1081

# Attachment MAPI property IDs
PR_ATTACH_LONG_FILENAME = 0x3707
PR_ATTACH_FILENAME = 0x3704
PR_ATTACH_EXTENSION = 0x3703
PR_ATTACH_CONTENT_ID = 0x3712
PR_ATTACH_METHOD = 0x3705
PR_ATTACH_MIME_TYPE = 0x3716
PR_DISPLAY_NAME = 0x3001
PR_ATTACH_SIZE = 0x0E20

# PR_ATTACH_METHOD values
ATTACH_BY_VALUE = 1
ATTACH_EMBEDDED_MSG = 5


class MIMEConverter:
    def __init__(
        self,
        hostname: Optional[str] = None,
        default_encoding: str = "cp1252",
    ):
        self.hostname = hostname or socket.gethostname() or "ost2maildir.local"
        self.default_encoding = default_encoding

    def convert(self, msg) -> tuple[EmailMessage | MIMEMultipart, str]:
        """
        Convert a pypff message to an RFC 2822 MIME message.

        Returns (mime_message, maildir_flags_string).
        """
        props = get_all_mapi_props(msg)
        flags_int = entry_as_int(props.get(PR_MESSAGE_FLAGS)) or 0
        maildir_flags = mapi_flags_to_maildir(flags_int, props)

        # Prefer transport_headers as the header source when available
        transport_headers = _safe_attr(msg, "transport_headers")

        # Build body MIME part(s)
        body_part = self._build_body(msg)

        # Build attachment parts
        attachment_parts = self._build_attachments(msg)

        # Assemble the message
        if attachment_parts:
            mime_msg = MIMEMultipart("mixed")
            mime_msg.attach(body_part)
            for att_part in attachment_parts:
                mime_msg.attach(att_part)
        else:
            mime_msg = body_part

        # Apply headers
        self._apply_headers(mime_msg, msg, props, transport_headers)

        # Decode any TNEF blobs hiding real attachments
        if mime_msg.is_multipart():
            process_tnef(mime_msg)

        return mime_msg, maildir_flags

    def _build_body(self, msg) -> MIMEText:
        """Build the body MIME part: HTML > plain text > RTF-converted."""
        html_bytes = _safe_attr(msg, "html_body")
        plain = _safe_attr(msg, "plain_text_body")
        rtf_bytes = _safe_attr(msg, "rtf_body")

        if html_bytes:
            charset = detect_html_charset(html_bytes)
            decoded = decode_body(html_bytes, charset)
            if decoded:
                return MIMEText(decoded, "html", "utf-8")

        if plain:
            decoded = decode_body(plain, self.default_encoding)
            if decoded:
                return MIMEText(decoded, "plain", "utf-8")

        if rtf_bytes:
            plain_from_rtf = _rtf_to_text(rtf_bytes)
            if plain_from_rtf:
                return MIMEText(plain_from_rtf, "plain", "utf-8")

        return MIMEText("", "plain", "utf-8")

    def _build_attachments(self, msg) -> list:
        """Build MIME parts for all attachments; skip inline body-only TNEF."""
        parts = []
        try:
            n = msg.number_of_attachments
        except Exception:
            return parts

        for i in range(n):
            try:
                att = msg.get_attachment(i)
                part = self._attachment_to_mime(att)
                if part is not None:
                    parts.append(part)
            except Exception as exc:
                log.debug("Skipping attachment %d: %s", i, exc)

        return parts

    def _attachment_to_mime(self, att) -> Optional[object]:
        """Convert a single pypff attachment to a MIME part."""
        att_props = _get_att_props(att)

        method = att_props.get("method", ATTACH_BY_VALUE)
        if method not in (ATTACH_BY_VALUE, ATTACH_EMBEDDED_MSG):
            log.debug("Skipping attachment with method %d", method)
            return None

        filename = (
            att_props.get("long_filename")
            or att_props.get("short_filename")
            or att_props.get("display_name")
            or "attachment"
        )
        ext = att_props.get("extension", "")
        if ext and not filename.endswith(ext):
            filename = filename + ext

        mime_type = att_props.get("mime_type") or mimetypes.guess_type(filename)[0] or "application/octet-stream"
        maintype, subtype = mime_type.split("/", 1)

        # Stream attachment data
        try:
            data = _read_attachment(att)
        except Exception as exc:
            log.debug("Cannot read attachment data for %r: %s", filename, exc)
            return None

        if not data:
            return None

        part = MIMEApplication(data, _subtype=subtype)
        safe_filename = filename.replace('"', "'")
        part["Content-Disposition"] = f'attachment; filename="{safe_filename}"'
        part.replace_header("Content-Type", f"{maintype}/{subtype}; name=\"{safe_filename}\"")

        content_id = att_props.get("content_id")
        if content_id:
            part["Content-ID"] = f"<{content_id}>"
            part.replace_header("Content-Disposition", f'inline; filename="{safe_filename}"')

        return part

    def _apply_headers(self, mime_msg, msg, props: dict, transport_headers: Optional[str]):
        """Set RFC 2822 headers on the assembled message."""
        # --- Date ---
        date = (
            _safe_attr(msg, "delivery_time")
            or _safe_attr(msg, "client_submit_time")
            or _safe_attr(msg, "creation_time")
        )
        if date:
            if date.tzinfo is None:
                date = date.replace(tzinfo=datetime.timezone.utc)
            mime_msg["Date"] = format_datetime(date, usegmt=True)
        else:
            mime_msg["Date"] = format_datetime(
                datetime.datetime.now(datetime.timezone.utc), usegmt=True
            )

        # --- From ---
        from_addr = _build_from(props)
        mime_msg["From"] = from_addr

        # --- To / CC ---
        # Prefer transport_headers for To/CC since they have real SMTP addresses
        to_hdr = cc_hdr = None
        if transport_headers:
            to_hdr, cc_hdr = _extract_to_cc_from_transport(transport_headers)

        if not to_hdr:
            display_to = entry_as_string(props.get(PR_DISPLAY_TO)) or ""
            if display_to:
                to_hdr = display_to
        if not cc_hdr:
            display_cc = entry_as_string(props.get(PR_DISPLAY_CC)) or ""
            if display_cc:
                cc_hdr = display_cc

        if to_hdr:
            mime_msg["To"] = to_hdr
        if cc_hdr:
            mime_msg["Cc"] = cc_hdr

        # --- Subject ---
        subject = clean_subject(_safe_attr(msg, "subject") or entry_as_string(props.get(PR_SUBJECT)) or "")
        mime_msg["Subject"] = subject

        # --- Message-ID ---
        mid = entry_as_string(props.get(PR_INTERNET_MESSAGE_ID))
        if not mid and transport_headers:
            mid = _extract_header_field(transport_headers, "Message-ID")
        if not mid:
            mid = synthetic_message_id(msg, self.hostname)
        mime_msg["Message-ID"] = mid

        # --- Threading ---
        in_reply_to = entry_as_string(props.get(PR_IN_REPLY_TO_ID))
        if not in_reply_to and transport_headers:
            in_reply_to = _extract_header_field(transport_headers, "In-Reply-To")
        if in_reply_to:
            mime_msg["In-Reply-To"] = in_reply_to

        references = entry_as_string(props.get(PR_INTERNET_REFERENCES))
        if not references and transport_headers:
            references = _extract_header_field(transport_headers, "References")
        if references:
            mime_msg["References"] = references

        # --- Importance / Priority ---
        importance = entry_as_int(props.get(PR_IMPORTANCE))
        if importance == 0:
            mime_msg["X-Priority"] = "5 (Lowest)"
            mime_msg["Importance"] = "low"
        elif importance == 2:
            mime_msg["X-Priority"] = "1 (Highest)"
            mime_msg["Importance"] = "high"

        # --- MIME version (if not already set) ---
        if not mime_msg.get("MIME-Version"):
            mime_msg["MIME-Version"] = "1.0"

        # --- X-ost2maildir provenance ---
        mime_msg["X-Converted-By"] = "ost2maildir/1.0"


# ── Module-level helpers ────────────────────────────────────────────────────

def _safe_attr(obj, attr: str):
    try:
        return getattr(obj, attr)
    except Exception:
        return None


def _rtf_to_text(rtf_bytes: bytes) -> Optional[str]:
    try:
        from striprtf.striprtf import rtf_to_text  # type: ignore
        text = rtf_to_text(rtf_bytes.decode("latin-1", errors="replace"))
        return text.strip() or None
    except Exception as exc:
        log.debug("RTF conversion failed: %s", exc)
        return None


def _build_from(props: dict) -> str:
    """Build From: header from MAPI sender properties."""
    name = entry_as_string(props.get(PR_SENDER_NAME)) or entry_as_string(props.get(PR_SENT_REPRESENTING_NAME)) or ""
    smtp = entry_as_string(props.get(PR_SENDER_SMTP_ADDRESS)) or ""
    if not smtp or not is_smtp_address(smtp):
        addr_raw = entry_as_string(props.get(PR_SENDER_EMAIL_ADDRESS)) or ""
        smtp = addr_raw if is_smtp_address(addr_raw) else ""
    if not smtp:
        smtp = "unknown@ost2maildir.local"
    if name:
        return formataddr((name, smtp))
    return smtp


def _get_att_props(att) -> dict:
    """Read key properties from an attachment's record_set."""
    result: dict = {}
    try:
        if att.number_of_record_sets < 1:
            return result
        rs = att.get_record_set(0)
        for i in range(rs.number_of_entries):
            e = rs.get_entry(i)
            et = e.entry_type
            try:
                if et == PR_ATTACH_LONG_FILENAME:
                    result["long_filename"] = e.data_as_string
                elif et == PR_ATTACH_FILENAME:
                    result["short_filename"] = e.data_as_string
                elif et == PR_ATTACH_EXTENSION:
                    result["extension"] = e.data_as_string
                elif et == PR_ATTACH_CONTENT_ID:
                    result["content_id"] = e.data_as_string
                elif et == PR_ATTACH_MIME_TYPE:
                    result["mime_type"] = e.data_as_string
                elif et == PR_DISPLAY_NAME:
                    result["display_name"] = e.data_as_string
                elif et == PR_ATTACH_METHOD:
                    try:
                        result["method"] = e.data_as_integer
                    except Exception:
                        raw = e.data
                        if raw and len(raw) >= 4:
                            import struct
                            result["method"] = struct.unpack_from("<I", raw)[0]
            except Exception:
                pass
    except Exception as exc:
        log.debug("Error reading attachment properties: %s", exc)
    return result


def _read_attachment(att, chunk_size: int = 65536) -> bytes:
    """Stream-read attachment data."""
    try:
        att.seek_offset(0, 0)
    except Exception:
        try:
            att.seek_offset(0)
        except Exception:
            pass
    chunks = []
    while True:
        try:
            chunk = att.read_buffer(chunk_size)
        except Exception:
            break
        if not chunk:
            break
        chunks.append(chunk)
    return b"".join(chunks)


def _extract_to_cc_from_transport(transport_headers: str) -> tuple[str, str]:
    """Extract To: and Cc: header values from a raw transport headers string."""
    to = _extract_header_field(transport_headers, "To") or ""
    cc = _extract_header_field(transport_headers, "Cc") or ""
    return to, cc


def _extract_header_field(headers: str, field: str) -> Optional[str]:
    """
    Extract a single header field value from a raw headers block.
    Handles multi-line (folded) headers.
    """
    import re
    pattern = re.compile(
        r"^" + re.escape(field) + r"\s*:\s*(.*?)(?=\n[^\s]|\Z)",
        re.IGNORECASE | re.MULTILINE | re.DOTALL,
    )
    m = pattern.search(headers)
    if m:
        # Unfold: replace \r\n + whitespace with a single space
        value = re.sub(r"\r?\n[ \t]+", " ", m.group(1)).strip()
        return value or None
    return None
