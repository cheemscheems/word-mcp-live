"""Periodic and on-demand document backup manager.

Usage::

    from word_document_server.utils.backup_manager import backup_manager

    # On-demand backup (call before destructive operations)
    result = backup_manager.create_backup("document.docx",
                                          note="before encrypt")

    # The background loop is started automatically by ``run_server()``.
"""

import asyncio
import datetime
import glob
import os
import shutil
import sys
import time


class BackupManager:
    """Manages automatic backups of Word documents.

    *Timer-based*: every 5 minutes, backed up files are snapshotted.
    *On-demand*: call ``create_backup()`` from any tool.

    Backups are stored in ``<source_dir>/<stem>_backup/``.
    When the number of backups exceeds ``MCP_MAX_BACKUPS`` (default 5),
    the oldest copies are automatically removed.

    This is a singleton — use ``backup_manager`` at module level.
    """

    INTERVAL = 300  # 5 minutes (seconds)
    MAX_BACKUPS = int(os.environ.get("MCP_MAX_BACKUPS", "5"))

    def __init__(self) -> None:
        # filepath (realpath) → last_track_timestamp
        self._tracked: dict[str, float] = {}
        self._task: asyncio.Task | None = None

    # ── Public API ───────────────────────────────────────────────────

    def track(self, filepath: str | None) -> None:
        """Register *filepath* for periodic backup tracking.

        Called automatically from ``get_file_lock()`` — every destructive
        tool call automatically tracks the file it is about to modify.
        """
        if not filepath:
            return
        real = os.path.realpath(filepath)
        self._tracked[real] = time.time()

    def create_backup(self, filepath: str, note: str = "") -> dict:
        """Create a single backup copy of *filepath*.

        Args:
            filepath: Path to the file to back up.
            note:     Short reason (e.g. ``"before encrypt"``).  Included
                      in the return dict for LLM to show the user.

        Returns:
            Dict with ``success``, ``path``, ``label``, and optionally
            ``warning`` (when old backups were removed) and ``note``.
        """
        real_path = os.path.realpath(filepath)
        if not os.path.exists(real_path):
            return {"success": False, "error": f"文件不存在: {real_path}"}

        src_dir = os.path.dirname(real_path)
        stem = os.path.splitext(os.path.basename(real_path))[0]
        backup_dir = os.path.join(src_dir, f"{stem}_backup")
        os.makedirs(backup_dir, exist_ok=True)

        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{stem}_backup_{ts}.docx"
        backup_path = os.path.join(backup_dir, backup_name)

        try:
            shutil.copy2(real_path, backup_path)
        except Exception as e:
            return {"success": False, "error": f"备份失败: {str(e)}"}

        # Enforce max backups
        deleted: list[str] = []
        if self.MAX_BACKUPS > 0:
            pattern = os.path.join(backup_dir, f"{stem}_backup_*.docx")
            existing = sorted(glob.glob(pattern))
            while len(existing) > self.MAX_BACKUPS:
                oldest = existing.pop(0)
                try:
                    os.remove(oldest)
                    deleted.append(os.path.basename(oldest))
                except Exception:
                    pass

        result: dict = {
            "success": True,
            "path": backup_path,
            "label": f"备份_{ts}",
        }
        if note:
            result["note"] = note
        if deleted:
            result["deleted_old"] = deleted
            result["warning"] = f"已删除旧备份：{', '.join(deleted)}"
        return result

    # ── Background loop ──────────────────────────────────────────────

    async def _periodic_loop(self) -> None:
        """Background task: backup tracked files every ``INTERVAL`` sec."""
        while True:
            await asyncio.sleep(self.INTERVAL)
            stale_threshold = time.time() - 1800  # 30 min inactivity
            for fpath in list(self._tracked.keys()):
                if not os.path.exists(fpath):
                    del self._tracked[fpath]
                    continue
                self.create_backup(fpath, note="auto")
                if self._tracked.get(fpath, 0) < stale_threshold:
                    del self._tracked[fpath]

    def start(self) -> None:
        """Launch the background backup loop.

        Called once from ``run_server()`` in *main.py*.
        Safe to call multiple times — only one loop runs.
        """
        if self._task is not None:
            return
        try:
            loop = asyncio.get_event_loop()
            self._task = loop.create_task(self._periodic_loop())
        except RuntimeError:
            pass  # no event loop yet

    def stop(self) -> None:
        """Cancel the background backup loop."""
        if self._task:
            self._task.cancel()
            self._task = None


# Singleton
backup_manager = BackupManager()
