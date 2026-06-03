"""Folder name mapping: OST/Outlook names → Maildir++ dot-path names."""

import re

# System/internal folders to skip entirely during iteration
SKIP_FOLDER_NAMES = {
    "common views",
    "finder",
    "views",
    "shortcuts",
    "~mapisp(internal)",
    "drizzle",
    "spam search folder 2",
    "shared data",
    "itemprocSearch",
    "itemprocssearch",
    "eventcheckpoints",
    "personmetadata",
    "conversation action settings",
    "quick step settings",
    "yammer root",
    "social activity notifications",
    "rss feeds",
    "externalcontacts",
    "peoplecentricconversation buddies",
    "organizational contacts",
    "gal contacts",
    "companies",
    "recipient cache",
    "skype for business contacts",
}

# Standard Outlook folder names → canonical Maildir++ subfolder names
# Keys are lowercased for case-insensitive matching
STANDARD_FOLDER_MAP = {
    "inbox": "",          # root Maildir (INBOX)
    "sent items": "Sent",
    "sent": "Sent",
    "deleted items": "Trash",
    "trash": "Trash",
    "drafts": "Drafts",
    "junk e-mail": "Junk",
    "junk email": "Junk",
    "junk": "Junk",
    "spam": "Junk",
    "outbox": "Outbox",
    "archive": "Archive",
    "calendar": "Calendar",
    "contacts": "Contacts",
    "tasks": "Tasks",
    "notes": "Notes",
    "sync issues": "SyncIssues",
    "clutter": "Clutter",
    "noise": "Noise",
    "conversation history": "ConversationHistory",
}

# Folders that contain items but should be placed under a parent
# e.g. Calendar sub-items (Birthdays, Canada holidays) go under .Calendar
CALENDAR_FOLDER_NAMES = {"birthdays", "canada holidays", "holidays"}


def sanitize_component(name: str) -> str:
    """Make a single folder name component safe for Maildir++ and filesystems."""
    if not name:
        return "unnamed"
    # Replace filesystem-unsafe characters
    name = re.sub(r'[/\\:*?"<>|]', "_", name)
    # Replace dots (Maildir++ hierarchy separator)
    name = name.replace(".", "_")
    # Collapse runs of spaces/underscores
    name = re.sub(r"[\s_]+", "_", name)
    name = name.strip("_")
    return name or "unnamed"


class FolderMapper:
    """Maps a list of OST folder path components to a Maildir++ folder name."""

    def __init__(self):
        # Track all seen folder names for generating subscriptions file
        self._seen: set[str] = set()

    def get_maildir_name(self, path_parts: list[str]) -> str:
        """
        Convert OST folder path to Maildir++ name.

        path_parts is the path relative to IPM_SUBTREE root, e.g.:
          ["Inbox"]                  → "" (root INBOX)
          ["Inbox", "Bulletins"]     → ".Bulletins"  (direct child of Inbox)
          ["Sent Items"]             → ".Sent"
          ["Work", "Projects", "X"] → ".Work.Projects.X"
        """
        if not path_parts:
            return ""

        # Normalise the top-level component
        top = path_parts[0].lower()
        rest = path_parts[1:]

        if top == "inbox":
            if not rest:
                return ""  # INBOX itself → root
            # Children of Inbox get placed at top level (same as Inbox)
            components = [sanitize_component(p) for p in rest]
        elif top in STANDARD_FOLDER_MAP:
            canonical = STANDARD_FOLDER_MAP[top]
            if not canonical:
                return ""
            components = [canonical] + [sanitize_component(p) for p in rest]
        else:
            components = [sanitize_component(path_parts[0])] + [sanitize_component(p) for p in rest]

        name = ".".join(components)
        if name:
            self._seen.add(name)
        return f".{name}" if name else ""

    def all_folders(self) -> list[str]:
        """Return sorted list of all folder names for the subscriptions file."""
        return sorted(self._seen)
