"""TNEF / winmail.dat decoding: expand hidden attachments from Exchange RTF messages."""

import logging
import mimetypes
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

log = logging.getLogger(__name__)


def process_tnef(mime_msg) -> bool:
    """
    Detect winmail.dat in mime_msg, decode it via tnefparse, and replace the
    winmail.dat attachment with the real attachments it contained.

    Returns True if TNEF was found and processed.
    """
    try:
        from tnefparse import TNEF  # type: ignore
    except ImportError:
        log.debug("tnefparse not available; skipping TNEF decode")
        return False

    # Only multipart messages can contain a winmail.dat attachment
    if not mime_msg.is_multipart():
        return False

    winmail_part = None
    winmail_idx = None
    payload = mime_msg.get_payload()

    for idx, part in enumerate(payload):
        ct = part.get_content_type()
        fn = part.get_filename() or ""
        if ct == "application/ms-tnef" or fn.lower() == "winmail.dat":
            winmail_part = part
            winmail_idx = idx
            break

    if winmail_part is None:
        return False

    tnef_data = winmail_part.get_payload(decode=True)
    if not tnef_data:
        return False

    try:
        tnef = TNEF(tnef_data)
    except Exception as exc:
        log.warning("TNEF decode failed; keeping winmail.dat as-is: %s", exc)
        return False

    # Remove the winmail.dat part
    payload.pop(winmail_idx)

    # Extract and add real attachments
    extracted = 0
    for att in tnef.attachments:
        try:
            filename = getattr(att, "long_filename", None) or getattr(att, "name", None) or "attachment"
            data = att.data
            if not data:
                continue
            mime_type, _ = mimetypes.guess_type(filename)
            maintype, subtype = (mime_type or "application/octet-stream").split("/", 1)
            part = MIMEApplication(data, _subtype=subtype)
            part["Content-Disposition"] = f'attachment; filename="{filename}"'
            part["Content-Type"] = f"{maintype}/{subtype}; name=\"{filename}\""
            payload.append(part)
            extracted += 1
        except Exception as exc:
            log.debug("Could not extract TNEF attachment: %s", exc)

    # If TNEF had an HTML or RTF body and the message body is empty, substitute it
    if hasattr(tnef, "htmlbody") and tnef.htmlbody:
        _replace_body(mime_msg, tnef.htmlbody, "text/html")
    elif hasattr(tnef, "body") and tnef.body:
        _replace_body(mime_msg, tnef.body, "text/plain")

    log.debug("TNEF decoded: %d real attachments extracted", extracted)
    return True


def _replace_body(mime_msg, body_data: bytes, content_type: str):
    """Replace the first text/* body part with new body_data."""
    payload = mime_msg.get_payload()
    for idx, part in enumerate(payload):
        if part.get_content_maintype() == "text":
            from email.mime.text import MIMEText
            maintype, subtype = content_type.split("/", 1)
            try:
                text = body_data.decode("utf-8") if isinstance(body_data, bytes) else body_data
            except UnicodeDecodeError:
                text = body_data.decode("latin-1", errors="replace")
            new_part = MIMEText(text, _subtype=subtype, _charset="utf-8")
            payload[idx] = new_part
            return
