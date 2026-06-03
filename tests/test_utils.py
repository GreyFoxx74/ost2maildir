"""Tests for utility functions."""

import datetime
import pytest
from ost2maildir.utils import (
    clean_subject,
    is_smtp_address,
    mapi_flags_to_maildir,
    synthetic_message_id,
    detect_html_charset,
    decode_body,
    filetime_to_datetime,
)


def test_clean_subject_strips_mapi_prefix():
    assert clean_subject("\x01\x05Re: Hello") == "Re: Hello"


def test_clean_subject_no_prefix():
    assert clean_subject("Hello World") == "Hello World"


def test_clean_subject_calendar_prefix():
    assert clean_subject("\x01\x01Team Standup") == "Team Standup"


def test_clean_subject_empty():
    assert clean_subject("") == ""
    assert clean_subject(None) == ""


def test_is_smtp_address_valid():
    assert is_smtp_address("user@example.com") is True
    assert is_smtp_address("First.Last@example.com") is True


def test_is_smtp_address_x500():
    assert is_smtp_address("/O=EXCHANGELABS/OU=Exchange Administrative Group/CN=RECIPIENTS/CN=abc") is False
    assert is_smtp_address("CN=John Smith") is False


def test_is_smtp_address_empty():
    assert is_smtp_address("") is False
    assert is_smtp_address(None) is False


def test_mapi_flags_read_message():
    # MSGFLAG_READ = 0x0001
    flags = mapi_flags_to_maildir(0x0001, {})
    assert "S" in flags


def test_mapi_flags_unread_message():
    flags = mapi_flags_to_maildir(0x0000, {})
    assert "S" not in flags


def test_mapi_flags_draft():
    flags = mapi_flags_to_maildir(0x0101, {})  # READ + UNSENT
    assert "D" in flags
    assert "S" in flags


def test_mapi_flags_sorted():
    flags = mapi_flags_to_maildir(0x0001, {})
    assert flags == "".join(sorted(flags))


def test_detect_html_charset_utf8():
    html = b'<meta charset="utf-8"><html>'
    assert detect_html_charset(html) == "utf-8"


def test_detect_html_charset_windows():
    html = b'<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">'
    assert "1252" in detect_html_charset(html)


def test_detect_html_charset_fallback():
    assert detect_html_charset(b"<html>no charset here</html>") == "utf-8"


def test_decode_body_bytes_utf8():
    result = decode_body("héllo".encode("utf-8"), "utf-8")
    assert result == "héllo"


def test_decode_body_already_str():
    assert decode_body("hello") == "hello"


def test_decode_body_empty():
    assert decode_body(None) is None
    assert decode_body(b"") is None


def test_filetime_to_datetime():
    # Known: 2024-01-01 00:00:00 UTC
    # Windows FILETIME for 2024-01-01T00:00:00Z
    ft = 133_485_408_000_000_000
    dt = filetime_to_datetime(ft)
    assert dt.year == 2024
    assert dt.month == 1
    assert dt.day == 1
    assert dt.tzinfo == datetime.timezone.utc
