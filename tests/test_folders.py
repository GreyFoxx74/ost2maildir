"""Tests for folder name mapping."""

import pytest
from ost2maildir.folders import FolderMapper, sanitize_component


def test_inbox_maps_to_root():
    fm = FolderMapper()
    assert fm.get_maildir_name(["Inbox"]) == ""


def test_sent_items():
    fm = FolderMapper()
    assert fm.get_maildir_name(["Sent Items"]) == ".Sent"


def test_deleted_items():
    fm = FolderMapper()
    assert fm.get_maildir_name(["Deleted Items"]) == ".Trash"


def test_drafts():
    fm = FolderMapper()
    assert fm.get_maildir_name(["Drafts"]) == ".Drafts"


def test_junk_email():
    fm = FolderMapper()
    assert fm.get_maildir_name(["Junk E-mail"]) == ".Junk"
    assert fm.get_maildir_name(["Junk Email"]) == ".Junk"


def test_inbox_subfolder():
    fm = FolderMapper()
    assert fm.get_maildir_name(["Inbox", "Bulletins"]) == ".Bulletins"


def test_custom_folder():
    fm = FolderMapper()
    result = fm.get_maildir_name(["Work"])
    assert result == ".Work"


def test_nested_custom_folder():
    fm = FolderMapper()
    result = fm.get_maildir_name(["Work", "Projects", "Alpha"])
    assert result == ".Work.Projects.Alpha"


def test_empty_path():
    fm = FolderMapper()
    assert fm.get_maildir_name([]) == ""


def test_folder_tracked_in_subscriptions():
    fm = FolderMapper()
    fm.get_maildir_name(["Sent Items"])
    fm.get_maildir_name(["Drafts"])
    folders = fm.all_folders()
    assert "Sent" in folders
    assert "Drafts" in folders


def test_sanitize_removes_dots():
    assert "." not in sanitize_component("file.name")


def test_sanitize_removes_slash():
    assert "/" not in sanitize_component("path/name")


def test_sanitize_handles_empty():
    assert sanitize_component("") == "unnamed"


def test_sanitize_collapses_spaces():
    result = sanitize_component("hello  world")
    assert "  " not in result
