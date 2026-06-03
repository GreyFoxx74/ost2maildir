"""Message-ID deduplication with optional persistence across runs."""

import json
import logging
import os
from pathlib import Path

log = logging.getLogger(__name__)

SEEN_FILE = ".ost2maildir_seen"


class Deduplicator:
    def __init__(self, output_dir: str, enabled: bool = True):
        self.enabled = enabled
        self._seen: set[str] = set()
        self._seen_path = Path(output_dir) / SEEN_FILE
        if enabled:
            self._load()

    def _load(self):
        if self._seen_path.exists():
            try:
                data = json.loads(self._seen_path.read_text())
                self._seen = set(data)
                log.info("Loaded %d previously seen Message-IDs", len(self._seen))
            except Exception as exc:
                log.warning("Could not load seen file: %s", exc)

    def is_duplicate(self, message_id: str) -> bool:
        if not self.enabled:
            return False
        if message_id in self._seen:
            return True
        self._seen.add(message_id)
        return False

    def save(self):
        if not self.enabled:
            return
        try:
            self._seen_path.write_text(json.dumps(list(self._seen)))
        except Exception as exc:
            log.warning("Could not save seen file: %s", exc)
