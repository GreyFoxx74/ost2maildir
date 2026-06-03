"""Calendar converter: IPM.Appointment → RFC 2822 message with .ics attachment."""

import datetime
import logging
import socket
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import format_datetime

from .utils import (
    clean_subject,
    entry_as_datetime,
    entry_as_int,
    entry_as_string,
    get_all_mapi_props,
    synthetic_message_id,
)

log = logging.getLogger(__name__)

# MAPI property IDs for appointments
PR_SUBJECT = 0x0037
PR_BODY = 0x1000
PR_CLIENT_SUBMIT_TIME = 0x0039
PR_MESSAGE_FLAGS = 0x0E07
PR_SENSITIVITY = 0x0036
PR_IMPORTANCE = 0x0017
PR_START_DATE = 0x0060
PR_END_DATE = 0x0061
PR_CREATION_TIME = 0x3007
PR_LAST_MOD_TIME = 0x3008


def convert_calendar(msg, hostname: str | None = None) -> tuple[MIMEMultipart, str]:
    """
    Convert an IPM.Appointment to an RFC 2822 message wrapping a .ics attachment.

    Returns (mime_message, maildir_flags).
    """
    if hostname is None:
        hostname = socket.gethostname() or "ost2maildir.local"

    props = get_all_mapi_props(msg)
    subject = clean_subject(
        getattr(msg, "subject", None) or entry_as_string(props.get(PR_SUBJECT)) or "Calendar Event"
    )

    ics_data = _build_ics(msg, props, hostname)

    outer = MIMEMultipart("mixed")
    outer["Subject"] = f"[Calendar] {subject}"
    outer["From"] = f"calendar@{hostname}"
    outer["Date"] = format_datetime(datetime.datetime.now(datetime.timezone.utc), usegmt=True)
    outer["Message-ID"] = synthetic_message_id(msg, hostname)
    outer["X-Converted-By"] = "ost2maildir/1.0"
    outer["MIME-Version"] = "1.0"

    body = MIMEText(f"Calendar event: {subject}", "plain", "utf-8")
    outer.attach(body)

    ics_part = MIMEText(ics_data, "calendar", "utf-8")
    ics_part["Content-Disposition"] = f'attachment; filename="event.ics"'
    outer.attach(ics_part)

    flags_int = entry_as_int(props.get(PR_MESSAGE_FLAGS)) or 0
    maildir_flags = "S" if flags_int & 0x0001 else ""

    return outer, maildir_flags


def _build_ics(msg, props: dict, hostname: str) -> str:
    """Build iCal VEVENT string from MAPI properties."""
    try:
        from icalendar import Calendar, Event  # type: ignore
    except ImportError:
        log.warning("icalendar not available; producing minimal .ics")
        return _minimal_ics(msg, props, hostname)

    cal = Calendar()
    cal.add("prodid", f"-//ost2maildir//EN")
    cal.add("version", "2.0")

    event = Event()

    # UID
    event.add("uid", synthetic_message_id(msg, hostname).strip("<>"))

    # Summary
    subject = clean_subject(
        getattr(msg, "subject", None) or entry_as_string(props.get(PR_SUBJECT)) or ""
    )
    event.add("summary", subject)

    # Description from body
    body = None
    try:
        body = getattr(msg, "plain_text_body", None)
    except Exception:
        pass
    if not body:
        body = entry_as_string(props.get(PR_BODY))
    if body:
        event.add("description", body[:2000])  # cap to avoid huge ICS

    # Start / End times
    start = entry_as_datetime(props.get(PR_START_DATE))
    end = entry_as_datetime(props.get(PR_END_DATE))

    if start:
        if start.tzinfo is None:
            start = start.replace(tzinfo=datetime.timezone.utc)
        event.add("dtstart", start)
    if end:
        if end.tzinfo is None:
            end = end.replace(tzinfo=datetime.timezone.utc)
        event.add("dtend", end)

    # Timestamps
    created = entry_as_datetime(props.get(PR_CREATION_TIME))
    if created:
        if created.tzinfo is None:
            created = created.replace(tzinfo=datetime.timezone.utc)
        event.add("created", created)

    # Sensitivity → CLASS
    sensitivity = entry_as_int(props.get(PR_SENSITIVITY)) or 0
    cls = {0: "PUBLIC", 1: "PERSONAL", 2: "PRIVATE", 3: "CONFIDENTIAL"}.get(sensitivity, "PUBLIC")
    event.add("class", cls)

    # Importance → PRIORITY
    importance = entry_as_int(props.get(PR_IMPORTANCE))
    if importance == 2:
        event.add("priority", 1)
    elif importance == 0:
        event.add("priority", 9)

    cal.add_component(event)
    return cal.to_ical().decode("utf-8", errors="replace")


def _minimal_ics(msg, props: dict, hostname: str) -> str:
    subject = clean_subject(
        getattr(msg, "subject", None) or entry_as_string(props.get(PR_SUBJECT)) or "Event"
    )
    uid = synthetic_message_id(msg, hostname).strip("<>")
    now = datetime.datetime.now(datetime.timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    start = entry_as_datetime(props.get(PR_START_DATE))
    dtstart = start.strftime("%Y%m%dT%H%M%SZ") if start else now
    return (
        "BEGIN:VCALENDAR\r\nVERSION:2.0\r\nPRODID:-//ost2maildir//EN\r\n"
        "BEGIN:VEVENT\r\n"
        f"UID:{uid}\r\n"
        f"DTSTAMP:{now}\r\n"
        f"DTSTART:{dtstart}\r\n"
        f"SUMMARY:{subject}\r\n"
        "END:VEVENT\r\nEND:VCALENDAR\r\n"
    )
