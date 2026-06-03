"""Maildir++ writer: create folder structure and write RFC 2822 messages atomically."""

import logging
import os
import re
import socket
import time
from pathlib import Path

log = logging.getLogger(__name__)

_pid = os.getpid()
_host = re.sub(r"[:/]", "_", socket.gethostname() or "localhost")
_seq = 0


def _next_seq() -> int:
    global _seq
    _seq += 1
    return _seq


class MaildirWriter:
    def __init__(self, output_dir: str):
        self.root = Path(output_dir)
        self._folders: set[str] = set()
        self._stats = {"written": 0, "errors": 0}
        self._ensure_folder(self.root)

    def write(
        self,
        maildir_name: str,
        mime_msg,
        maildir_flags: str = "",
        timestamp: float | None = None,
    ) -> bool:
        """
        Write a MIME message to the correct Maildir++ folder.

        maildir_name: "" for INBOX, ".Sent" etc for subfolders
        maildir_flags: string of sorted flag letters, e.g. "RS"
        timestamp: Unix epoch float for filename (defaults to now)

        Returns True on success.
        """
        if maildir_name:
            folder_dir = self.root / maildir_name
        else:
            folder_dir = self.root

        self._ensure_folder(folder_dir)

        if maildir_name:
            self._folders.add(maildir_name.lstrip("."))

        ts = int(timestamp or time.time())
        seq = _next_seq()
        safe_flags = "".join(sorted(set(maildir_flags)))
        filename = f"{ts}.{seq}_{_pid}.{_host}:2,{safe_flags}"

        tmp_path = folder_dir / "tmp" / filename
        cur_path = folder_dir / "cur" / filename

        try:
            msg_bytes = _serialize_message(mime_msg)
            tmp_path.write_bytes(msg_bytes)
            tmp_path.rename(cur_path)
            self._stats["written"] += 1
            return True
        except Exception as exc:
            log.error("Failed to write message %r: %s", filename, exc)
            self._stats["errors"] += 1
            try:
                tmp_path.unlink(missing_ok=True)
            except Exception:
                pass
            return False

    def finalize(self):
        """Write subscriptions file and return stats."""
        subs_path = self.root / "subscriptions"
        try:
            subs_path.write_text("\n".join(sorted(self._folders)) + "\n")
            log.info("Wrote subscriptions file: %d folders", len(self._folders))
        except Exception as exc:
            log.warning("Could not write subscriptions: %s", exc)

        log.info(
            "Maildir complete — written: %d, errors: %d",
            self._stats["written"],
            self._stats["errors"],
        )
        return dict(self._stats)

    def _ensure_folder(self, folder_dir: Path):
        for sub in ("new", "cur", "tmp"):
            (folder_dir / sub).mkdir(parents=True, exist_ok=True)


def _serialize_message(mime_msg) -> bytes:
    """Serialize a MIME message to bytes."""
    from email import generator
    import io

    buf = io.BytesIO()
    try:
        gen = generator.BytesGenerator(buf, mangle_from_=False)
        gen.flatten(mime_msg)
        result = buf.getvalue()
    except Exception:
        result = mime_msg.as_bytes()

    # Ensure \r\n line endings (IMAP / Maildir convention)
    if b"\r\n" not in result[:200] and b"\n" in result:
        result = result.replace(b"\n", b"\r\n")
    return result
