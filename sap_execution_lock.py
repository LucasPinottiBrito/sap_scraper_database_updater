from __future__ import annotations

from datetime import datetime
import logging
import os
from pathlib import Path
from typing import BinaryIO

try:
    import msvcrt
except ImportError:  # pragma: no cover - usado apenas fora do Windows.
    msvcrt = None

try:
    import fcntl
except ImportError:  # pragma: no cover - usado apenas fora de Unix.
    fcntl = None


logger = logging.getLogger(__name__)

BASE_DIR = Path(__file__).resolve().parent
DEFAULT_LOCK_FILE = BASE_DIR / ".sap_execution.lock"
LOCK_FILE_ENV = "SAP_EXECUTION_LOCK_FILE"


class SapExecutionLockBusy(RuntimeError):
    """Indica que outra execucao SAP ja esta em andamento."""


def get_sap_execution_lock_path() -> Path:
    configured_path = os.getenv(LOCK_FILE_ENV, "").strip()
    if configured_path:
        return Path(configured_path).expanduser().resolve()
    return DEFAULT_LOCK_FILE


class SapExecutionLock:
    """Lock cross-process para impedir duas automacoes SAP simultaneas."""

    def __init__(self, operation: str = "sap", lock_file: str | Path | None = None) -> None:
        self.operation = operation
        self.lock_file = Path(lock_file).expanduser().resolve() if lock_file else get_sap_execution_lock_path()
        self._handle: BinaryIO | None = None

    def __enter__(self) -> "SapExecutionLock":
        self.acquire()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.release()

    def acquire(self) -> None:
        self.lock_file.parent.mkdir(parents=True, exist_ok=True)
        handle = self.lock_file.open("a+b")
        self._ensure_lock_byte(handle)

        try:
            self._lock_handle(handle)
        except OSError as exc:
            handle.close()
            raise SapExecutionLockBusy(
                f"SAP ocupado por outra execucao. Lock: {self.lock_file}"
            ) from exc

        self._handle = handle
        self._write_metadata(handle)
        logger.info("Lock SAP adquirido para %s em %s", self.operation, self.lock_file)

    def release(self) -> None:
        if not self._handle:
            return

        handle = self._handle
        self._handle = None

        try:
            self._unlock_handle(handle)
            logger.info("Lock SAP liberado para %s em %s", self.operation, self.lock_file)
        finally:
            handle.close()

    @staticmethod
    def _ensure_lock_byte(handle: BinaryIO) -> None:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()
        handle.seek(0)

    @staticmethod
    def _lock_handle(handle: BinaryIO) -> None:
        handle.seek(0)
        if msvcrt:
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            return
        if fcntl:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return
        raise OSError("Plataforma sem suporte a lock de arquivo")

    @staticmethod
    def _unlock_handle(handle: BinaryIO) -> None:
        handle.seek(0)
        if msvcrt:
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
            return
        if fcntl:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)

    def _write_metadata(self, handle: BinaryIO) -> None:
        metadata = (
            f"pid={os.getpid()}\n"
            f"operation={self.operation}\n"
            f"acquired_at={datetime.now().isoformat()}\n"
        )
        handle.seek(0)
        handle.truncate()
        handle.write(metadata.encode("utf-8"))
        handle.flush()
