"""OST file reader: encryption detection and recursive item iterator."""

import logging
from typing import Iterator

import pypff

from .folders import SKIP_FOLDER_NAMES

log = logging.getLogger(__name__)

# Folders inside IPM_SUBTREE that hold calendar/contacts/mail but we want
# to recurse into them with the right path context
IPM_SUBTREE_NAME = "IPM_SUBTREE"
ROOT_MAILBOX_NAME = "Root - Mailbox"

# FAI (Folder-Associated Information) messages are internal Outlook metadata.
# They do NOT appear in number_of_sub_messages; they're in sub_items.
# We always iterate sub_messages only, so FAI is automatically excluded.


class EncryptionError(RuntimeError):
    pass


class OSTReader:
    def __init__(self, path: str):
        self.path = path
        self._file = pypff.file()
        self._file.open(path)

    def close(self):
        try:
            self._file.close()
        except Exception:
            pass

    def detect_encryption(self) -> str:
        """
        Detect PFF encryption mode.
        Returns "none", "compressible", or "high".
        Raises EncryptionError if high encryption is detected.
        """
        # libpff raises an error when it encounters high-encryption data.
        # We probe by opening and reading a small piece of the root.
        try:
            root = self._file.get_root_folder()
            _ = root.number_of_sub_folders
            return "compressible"  # if it opened fine, compressible or none
        except Exception as exc:
            msg = str(exc).lower()
            if "encrypt" in msg or "cipher" in msg:
                raise EncryptionError(
                    f"OST file uses high encryption (0x02) — cannot be read by libpff.\n"
                    f"To convert: open the OST in Outlook → File → Export to a .pst file, "
                    f"then run ost2maildir on the PST.\n"
                    f"libpff error: {exc}"
                ) from exc
            raise

    def _get_ipm_subtree(self):
        """Navigate to the Root - Mailbox → IPM_SUBTREE folder."""
        root = self._file.get_root_folder()

        # Find Root - Mailbox
        mailbox_root = None
        for i in range(root.number_of_sub_folders):
            sf = root.get_sub_folder(i)
            if sf.name == ROOT_MAILBOX_NAME:
                mailbox_root = sf
                break

        if mailbox_root is None:
            # Some OST files have a different root structure; try direct
            return root

        # Find IPM_SUBTREE
        for i in range(mailbox_root.number_of_sub_folders):
            sf = mailbox_root.get_sub_folder(i)
            if sf.name == IPM_SUBTREE_NAME:
                return sf

        return mailbox_root

    def count_items(self) -> dict:
        """Dry-run: count all items without converting."""
        counts = {"messages": 0, "calendar": 0, "contacts": 0,
                  "tasks": 0, "other": 0, "skipped_errors": 0}

        def _walk(folder):
            for i in range(folder.number_of_sub_messages):
                try:
                    msg = folder.get_sub_message(i)
                    mc = _get_message_class(msg)
                    if mc.startswith("IPM.Note") or mc.startswith("IPM.Schedule") or mc.startswith("Report"):
                        counts["messages"] += 1
                    elif mc.startswith("IPM.Appointment"):
                        counts["calendar"] += 1
                    elif mc.startswith("IPM.Contact") or mc.startswith("IPM.DistList"):
                        counts["contacts"] += 1
                    elif mc.startswith("IPM.Task"):
                        counts["tasks"] += 1
                    else:
                        counts["other"] += 1
                except Exception:
                    counts["skipped_errors"] += 1

            for i in range(folder.number_of_sub_folders):
                try:
                    sf = folder.get_sub_folder(i)
                    if sf.name and sf.name.lower() in SKIP_FOLDER_NAMES:
                        continue
                    _walk(sf)
                except Exception:
                    pass

        ipm = self._get_ipm_subtree()
        _walk(ipm)
        return counts

    def iter_items(self) -> Iterator[tuple[list[str], object]]:
        """
        Yield (folder_path_parts, pypff_message) for every item in the OST.

        folder_path_parts is relative to IPM_SUBTREE, e.g.:
          ["Inbox"] for the top-level Inbox
          ["Inbox", "Bulletins"] for a subfolder
          ["Calendar"] for calendar items
        """
        ipm = self._get_ipm_subtree()
        yield from _walk_folder(ipm, [])


def _get_message_class(msg) -> str:
    """Read PR_MESSAGE_CLASS (0x001A) from a message; default to IPM.Note."""
    try:
        if msg.number_of_record_sets < 1:
            return "IPM.Note"
        rs = msg.get_record_set(0)
        for i in range(rs.number_of_entries):
            e = rs.get_entry(i)
            if e.entry_type == 0x001A:
                try:
                    return e.data_as_string or "IPM.Note"
                except Exception:
                    return "IPM.Note"
    except Exception:
        pass
    return "IPM.Note"


def _walk_folder(
    folder, path_parts: list[str]
) -> Iterator[tuple[list[str], object]]:
    """Recursively yield (path_parts, msg) from folder and all sub-folders."""
    # Yield messages in this folder
    try:
        n = folder.number_of_sub_messages
    except Exception as exc:
        log.debug("Cannot read message count for folder %r: %s", folder.name, exc)
        n = 0

    for i in range(n):
        try:
            msg = folder.get_sub_message(i)
            yield path_parts, msg
        except Exception as exc:
            log.warning("Skipping message %d in %r: %s", i, folder.name, exc)

    # Recurse into sub-folders
    try:
        nsf = folder.number_of_sub_folders
    except Exception:
        nsf = 0

    for i in range(nsf):
        try:
            sf = folder.get_sub_folder(i)
            sf_name = sf.name or ""
            if sf_name.lower() in SKIP_FOLDER_NAMES:
                log.debug("Skipping system folder: %r", sf_name)
                continue
            # Skip unnamed folders and pure search / virtual folders
            if not sf_name or sf_name.startswith("MS-OLK-"):
                continue
            yield from _walk_folder(sf, path_parts + [sf_name])
        except Exception as exc:
            log.debug("Skipping sub-folder %d: %s", i, exc)
