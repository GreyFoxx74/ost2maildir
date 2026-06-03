"""Contact converter: IPM.Contact → RFC 2822 message with .vcf attachment."""

import datetime
import logging
import socket
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import format_datetime

from .utils import (
    clean_subject,
    entry_as_int,
    entry_as_string,
    get_all_mapi_props,
    is_smtp_address,
    synthetic_message_id,
)

log = logging.getLogger(__name__)

# Contact MAPI property IDs
PR_GIVEN_NAME = 0x3A06
PR_SURNAME = 0x3A11
PR_COMPANY_NAME = 0x3A16
PR_JOB_TITLE = 0x3A17
PR_DISPLAY_NAME = 0x3001
PR_BIZ_PHONE = 0x3A08
PR_HOME_PHONE = 0x3A09
PR_MOBILE_PHONE = 0x3A1C
PR_EMAIL_ADDRESS = 0x3003
PR_ADDRTYPE = 0x3002
PR_MESSAGE_FLAGS = 0x0E07
PR_BODY = 0x1000


def convert_contact(msg, hostname: str | None = None) -> tuple[MIMEMultipart, str]:
    """
    Convert an IPM.Contact to an RFC 2822 message wrapping a .vcf attachment.

    Returns (mime_message, maildir_flags).
    """
    if hostname is None:
        hostname = socket.gethostname() or "ost2maildir.local"

    props = get_all_mapi_props(msg)

    given = entry_as_string(props.get(PR_GIVEN_NAME)) or ""
    surname = entry_as_string(props.get(PR_SURNAME)) or ""
    display = entry_as_string(props.get(PR_DISPLAY_NAME)) or ""
    full_name = f"{given} {surname}".strip() or display or "Unknown Contact"

    vcf_data = _build_vcf(props, full_name)

    outer = MIMEMultipart("mixed")
    outer["Subject"] = f"[Contact] {full_name}"
    outer["From"] = f"contacts@{hostname}"
    outer["Date"] = format_datetime(datetime.datetime.now(datetime.timezone.utc), usegmt=True)
    outer["Message-ID"] = synthetic_message_id(msg, hostname)
    outer["X-Converted-By"] = "ost2maildir/1.0"
    outer["MIME-Version"] = "1.0"

    body = MIMEText(f"Contact: {full_name}", "plain", "utf-8")
    outer.attach(body)

    vcf_part = MIMEText(vcf_data, "x-vcard", "utf-8")
    vcf_part.replace_header("Content-Type", 'text/vcard; charset="utf-8"')
    vcf_part["Content-Disposition"] = f'attachment; filename="contact.vcf"'
    outer.attach(vcf_part)

    flags_int = entry_as_int(props.get(PR_MESSAGE_FLAGS)) or 0
    maildir_flags = "S" if flags_int & 0x0001 else ""

    return outer, maildir_flags


def _build_vcf(props: dict, full_name: str) -> str:
    """Build a vCard 3.0 string from MAPI properties."""
    lines = ["BEGIN:VCARD", "VERSION:3.0"]

    given = entry_as_string(props.get(PR_GIVEN_NAME)) or ""
    surname = entry_as_string(props.get(PR_SURNAME)) or ""
    lines.append(f"N:{_vcard_escape(surname)};{_vcard_escape(given)};;;")
    lines.append(f"FN:{_vcard_escape(full_name)}")

    company = entry_as_string(props.get(PR_COMPANY_NAME))
    if company:
        lines.append(f"ORG:{_vcard_escape(company)}")

    title = entry_as_string(props.get(PR_JOB_TITLE))
    if title:
        lines.append(f"TITLE:{_vcard_escape(title)}")

    biz = entry_as_string(props.get(PR_BIZ_PHONE))
    if biz:
        lines.append(f"TEL;TYPE=WORK,VOICE:{_vcard_escape(biz)}")

    home = entry_as_string(props.get(PR_HOME_PHONE))
    if home:
        lines.append(f"TEL;TYPE=HOME,VOICE:{_vcard_escape(home)}")

    mobile = entry_as_string(props.get(PR_MOBILE_PHONE))
    if mobile:
        lines.append(f"TEL;TYPE=CELL,VOICE:{_vcard_escape(mobile)}")

    # Email address — only use if SMTP
    email = entry_as_string(props.get(PR_EMAIL_ADDRESS))
    addrtype = entry_as_string(props.get(PR_ADDRTYPE)) or ""
    if email and (addrtype.upper() == "SMTP" or is_smtp_address(email)):
        lines.append(f"EMAIL;TYPE=INTERNET:{_vcard_escape(email)}")

    lines.append("END:VCARD")
    return "\r\n".join(lines) + "\r\n"


def _vcard_escape(s: str) -> str:
    """Escape special vCard characters."""
    if not s:
        return ""
    return s.replace("\\", "\\\\").replace(",", "\\,").replace(";", "\\;").replace("\n", "\\n")
