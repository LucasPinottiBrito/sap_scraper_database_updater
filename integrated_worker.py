from __future__ import annotations

import logging
import os
import time
from typing import Callable

from action_processor import ActionQueueProcessor, QueuedAction, configure_logging, get_poll_interval_seconds
from sap_execution_lock import SapExecutionLock, SapExecutionLockBusy


logger = logging.getLogger(__name__)

DEFAULT_FULL_UPDATE_INTERVAL_SECONDS = 1800
DEFAULT_FULL_UPDATE_RETRY_SECONDS = 300
FULL_UPDATE_INTERVAL_ENV = "SAP_FULL_UPDATE_INTERVAL_SECONDS"
FULL_UPDATE_RETRY_ENV = "SAP_FULL_UPDATE_RETRY_SECONDS"


def get_positive_int_env(env_name: str, default_value: int) -> int:
    raw_value = os.getenv(env_name)
    if not raw_value:
        return default_value

    try:
        value = int(raw_value)
    except ValueError:
        logger.warning("%s invalido (%s). Usando %s segundos.", env_name, raw_value, default_value)
        return default_value

    if value <= 0:
        logger.warning("%s deve ser maior que zero. Usando %s segundos.", env_name, default_value)
        return default_value

    return value


def run_full_update_without_lock() -> None:
    from run_updates import run_full_update

    run_full_update(acquire_lock=False)


class IntegratedSapWorker:
    def __init__(
        self,
        queue_processor: ActionQueueProcessor | None = None,
        full_update_runner: Callable[[], None] | None = None,
        lock_factory: Callable[[str], SapExecutionLock] | None = None,
        queue_poll_interval_seconds: int | None = None,
        full_update_interval_seconds: int | None = None,
        full_update_retry_seconds: int | None = None,
        sleep_func: Callable[[float], None] = time.sleep,
        clock: Callable[[], float] = time.monotonic,
    ) -> None:
        self.queue_processor = queue_processor or ActionQueueProcessor()
        self.full_update_runner = full_update_runner or run_full_update_without_lock
        self.lock_factory = lock_factory or (lambda operation: SapExecutionLock(operation=operation))
        self.queue_poll_interval_seconds = (
            queue_poll_interval_seconds
            if queue_poll_interval_seconds is not None
            else get_poll_interval_seconds()
        )
        self.full_update_interval_seconds = (
            full_update_interval_seconds
            if full_update_interval_seconds is not None
            else get_positive_int_env(FULL_UPDATE_INTERVAL_ENV, DEFAULT_FULL_UPDATE_INTERVAL_SECONDS)
        )
        self.full_update_retry_seconds = (
            full_update_retry_seconds
            if full_update_retry_seconds is not None
            else get_positive_int_env(FULL_UPDATE_RETRY_ENV, DEFAULT_FULL_UPDATE_RETRY_SECONDS)
        )
        self.sleep_func = sleep_func
        self.clock = clock
        self.next_full_update_at = 0.0

        self._normalize_intervals()

    def run_forever(self) -> None:
        logger.info(
            "Worker SAP integrado iniciado. Fila a cada %s segundos; atualizacao a cada %s segundos.",
            self.queue_poll_interval_seconds,
            self.full_update_interval_seconds,
        )

        while True:
            try:
                self.run_once()
            except KeyboardInterrupt:
                logger.info("Worker SAP integrado interrompido pelo usuario.")
                break
            except Exception:
                logger.exception("Erro inesperado no worker SAP integrado. O worker continuara em execucao.")

            try:
                self.sleep_func(self.queue_poll_interval_seconds)
            except KeyboardInterrupt:
                logger.info("Worker SAP integrado interrompido pelo usuario.")
                break

    def run_once(self) -> None:
        pending_actions = self._fetch_pending_actions()
        if pending_actions is None:
            return

        if pending_actions:
            queue_processed = self._process_queue_with_lock(pending_actions)
            if not queue_processed:
                return

        if self._is_full_update_due():
            self._run_full_update_with_lock()

    def _fetch_pending_actions(self) -> list[QueuedAction] | None:
        try:
            pending_actions = self.queue_processor.buscar_acoes_pendentes()
        except Exception:
            logger.exception("Erro ao buscar acoes pendentes. O worker tentara novamente no proximo ciclo.")
            return None

        logger.info("Fila SAP consultada: %s acao(oes) pendente(s).", len(pending_actions))
        return pending_actions

    def _process_queue_with_lock(self, pending_actions: list[QueuedAction]) -> bool:
        try:
            with self.lock_factory("fila SAP"):
                report = self.queue_processor.process_actions(pending_actions)
        except SapExecutionLockBusy:
            logger.warning(
                "SAP ocupado. %s acao(oes) da fila aguardarao o proximo ciclo.",
                len(pending_actions),
            )
            return False

        logger.info("Fila SAP finalizada: %s", report.get("message"))
        return True

    def _is_full_update_due(self) -> bool:
        return self.clock() >= self.next_full_update_at

    def _run_full_update_with_lock(self) -> bool:
        try:
            with self.lock_factory("atualizacao de bases SAP"):
                self.full_update_runner()
        except SapExecutionLockBusy:
            logger.warning("SAP ocupado. Atualizacao de bases aguardara o proximo ciclo.")
            return False
        except Exception:
            self.next_full_update_at = self.clock() + self.full_update_retry_seconds
            logger.exception(
                "Erro ao executar run_full_update(). Nova tentativa em %s segundos.",
                self.full_update_retry_seconds,
            )
            return False

        self.next_full_update_at = self.clock() + self.full_update_interval_seconds
        logger.info(
            "Atualizacao de bases concluida. Proxima tentativa em %s segundos.",
            self.full_update_interval_seconds,
        )
        return True

    def _normalize_intervals(self) -> None:
        self.queue_poll_interval_seconds = self._positive_or_default(
            self.queue_poll_interval_seconds,
            get_poll_interval_seconds(),
            "Intervalo da fila SAP",
        )
        self.full_update_interval_seconds = self._positive_or_default(
            self.full_update_interval_seconds,
            DEFAULT_FULL_UPDATE_INTERVAL_SECONDS,
            "Intervalo da atualizacao de bases",
        )
        self.full_update_retry_seconds = self._positive_or_default(
            self.full_update_retry_seconds,
            DEFAULT_FULL_UPDATE_RETRY_SECONDS,
            "Retry da atualizacao de bases",
        )

    @staticmethod
    def _positive_or_default(value: int, default_value: int, label: str) -> int:
        if value > 0:
            return value
        logger.warning("%s deve ser maior que zero. Usando %s segundos.", label, default_value)
        return default_value


def run_integrated_worker() -> None:
    configure_logging()
    IntegratedSapWorker().run_forever()


if __name__ == "__main__":
    run_integrated_worker()
