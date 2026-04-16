"""
Processador da fila de acoes SAP.

Fluxo:
1. busca acoes pendentes no backend;
2. valida os dados do registro da fila;
3. abre a nota na IW52;
4. executa a acao solicitada usando os campos da fila;
5. registra o resultado no backend.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import StrEnum
import logging
import os
from pathlib import Path
from typing import Any, Callable, Mapping, Optional

from dotenv import load_dotenv

from backend_client import BackendApiClient
from lib.screen.SapLogonScreen import SapLogonScreen

try:
    import pythoncom
except ImportError:  # pragma: no cover - pywin32 existe no ambiente Windows/SAP.
    pythoncom = None


logger = logging.getLogger(__name__)
load_dotenv(Path(__file__).resolve().parent / ".env")


class QueueAction(StrEnum):
    PROCEDER = "proceder"
    IMPROCEDER = "improceder"
    TRANSFERIR_SEMCI = "transferir_semci"
    TRANSFERIR_MMGD = "transferir_mmgd"
    ENCERRAR = "encerrar"


REGION_SAP_CONFIG: Mapping[str, Mapping[str, str]] = {
    "SP": {"ambiente": "EP1", "password_env": "SAP_PASSWORD_EP1"},
    "ES": {"ambiente": "EP2", "password_env": "SAP_PASSWORD_EP2"},
}


class ActionValidationError(ValueError):
    """Erro de entrada da fila antes de tentar qualquer automacao SAP."""


class SapExecutionError(RuntimeError):
    """Erro ao abrir a nota ou executar a acao no SAP."""


@dataclass(frozen=True)
class QueuedAction:
    id: int
    registration: str
    action: str
    note: str
    region: str
    created_at: datetime
    status: str = "pendente"
    texto_resultado: Optional[str] = None
    texto_medida: Optional[str] = None

    @classmethod
    def from_api(cls, payload: Mapping[str, Any]) -> "QueuedAction":
        return cls(
            id=int(payload["id"]),
            registration=str(payload.get("registration") or ""),
            action=str(payload.get("action") or ""),
            note=str(payload.get("note") or ""),
            region=str(payload.get("region") or ""),
            created_at=parse_datetime(payload.get("created_at")),
            status=str(payload.get("status") or "pendente"),
            texto_resultado=payload.get("texto_resultado"),
            texto_medida=payload.get("texto_medida"),
        )


@dataclass
class ExecutionResult:
    action_id: int
    success: bool
    error_message: Optional[str] = None
    timestamp: datetime = field(default_factory=datetime.now)

    def to_dict(self) -> dict[str, Any]:
        return {
            "action_id": self.action_id,
            "success": self.success,
            "error_message": self.error_message,
            "timestamp": self.timestamp.isoformat(),
        }


def parse_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, str) and value:
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00"))
        except ValueError:
            logger.debug("Nao foi possivel converter data da fila: %s", value)
    return datetime.now()


def require_text(value: Optional[str], field_name: str, action: QueuedAction) -> str:
    if value is None or value.strip() == "":
        raise ActionValidationError(
            f"Campo '{field_name}' e obrigatorio para a acao '{action.action}' "
            f"da fila {action.id}"
        )
    return value


class ComApartment:
    """Inicializa COM apenas durante a automacao SAP."""

    def __init__(self) -> None:
        self.initialized = False

    def __enter__(self):
        if pythoncom:
            pythoncom.CoInitialize()
            self.initialized = True
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        if self.initialized and pythoncom:
            pythoncom.CoUninitialize()


class SapIw52ActionRunner:
    """Abre a nota na IW52 e executa a acao pedida pela fila."""

    def __init__(
        self,
        logon_factory: Callable[[], SapLogonScreen] | None = None,
        env: Mapping[str, str] | None = None,
        region_config: Mapping[str, Mapping[str, str]] = REGION_SAP_CONFIG,
    ) -> None:
        self.logon_factory = logon_factory or SapLogonScreen
        self.env = env if env is not None else os.environ
        self.region_config = region_config

    def executar_acao_da_fila(self, action: QueuedAction) -> None:
        sap_config = self._get_region_config(action.region)
        sap_user, sap_password = self._get_credentials(action.region, sap_config)

        login_screen = None
        iw52_screen = None
        note_screen = None

        logger.info(
            "[ACTION_%s] Abrindo IW52/%s para nota %s",
            action.id,
            sap_config["ambiente"],
            action.note,
        )

        with ComApartment():
            try:
                logon = self.logon_factory()
                login_screen = logon.loadSystem(action.region, sap_config["ambiente"])
                home = login_screen.login(
                    sap_user,
                    sap_password,
                    action.region,
                    sap_config["ambiente"],
                )
                iw52_screen = home.openTransaction("iw52")
                note_screen = iw52_screen.openNote(action.note)

                self._executar_na_tela_da_nota(note_screen, action)
            except ActionValidationError:
                raise
            except Exception as exc:
                raise SapExecutionError(str(exc)) from exc
            finally:
                self._safe_back(note_screen, action)
                self._safe_close_iw52(iw52_screen, login_screen, action)

    def _executar_na_tela_da_nota(self, note_screen, action: QueuedAction) -> None:
        if action.action == QueueAction.PROCEDER.value:
            self.executar_proceder(note_screen, action)
            return

        if action.action == QueueAction.IMPROCEDER.value:
            self.executar_improceder(note_screen, action)
            return

        if action.action == QueueAction.TRANSFERIR_SEMCI.value:
            self.executar_transferir_para_sem_ci(note_screen)
            return

        if action.action == QueueAction.TRANSFERIR_MMGD.value:
            self.executar_transferir_para_mmgd(note_screen)
            return

        if action.action == QueueAction.ENCERRAR.value:
            self.executar_encerrar(note_screen)
            return

        raise ActionValidationError(f"Tipo de acao invalido: {action.action}")

    def executar_proceder(self, note_screen, action: QueuedAction) -> None:
        texto_resultado = require_text(action.texto_resultado, "texto_resultado", action)
        texto_medida = require_text(action.texto_medida, "texto_medida", action)

        logger.info("[ACTION_%s] Procedendo nota %s", action.id, action.note)
        note_screen.proceder(
            texto_resultado=texto_resultado,
            texto_medida=texto_medida,
        )

    def executar_improceder(self, note_screen, action: QueuedAction) -> None:
        texto_resultado = require_text(action.texto_resultado, "texto_resultado", action)
        texto_medida = require_text(action.texto_medida, "texto_medida", action)

        logger.info("[ACTION_%s] Improcedendo nota %s", action.id, action.note)
        note_screen.improceder(
            texto_resultado=texto_resultado,
            texto_medida=texto_medida,
        )

    @staticmethod
    def executar_transferir_para_sem_ci(note_screen) -> None:
        note_screen.transferir_para_sem_ci()

    @staticmethod
    def executar_transferir_para_mmgd(note_screen) -> None:
        note_screen.transferir_para_mmgd()

    @staticmethod
    def executar_encerrar(note_screen) -> None:
        note_screen.encerrar_nota()

    def _get_region_config(self, region: str) -> Mapping[str, str]:
        sap_config = self.region_config.get(region)
        if not sap_config:
            raise ActionValidationError(f"Regiao sem configuracao SAP: {region}")
        return sap_config

    def _get_credentials(self, region: str, sap_config: Mapping[str, str]) -> tuple[str, str]:
        sap_user = self.env.get("SAP_USER", "")
        sap_password = self.env.get(sap_config["password_env"], "")
        if not sap_user or not sap_password:
            raise ActionValidationError(f"Credenciais SAP ausentes para regiao {region}")
        return sap_user, sap_password

    @staticmethod
    def _safe_back(note_screen, action: QueuedAction) -> None:
        if not note_screen:
            return
        try:
            note_screen.back()
        except Exception:
            logger.debug(
                "[ACTION_%s] Nao foi possivel voltar da tela da nota",
                action.id,
                exc_info=True,
            )

    @staticmethod
    def _safe_close_iw52(iw52_screen, login_screen, action: QueuedAction) -> None:
        if iw52_screen:
            try:
                iw52_screen.close()
                return
            except Exception:
                logger.debug(
                    "[ACTION_%s] Nao foi possivel fechar IW52",
                    action.id,
                    exc_info=True,
                )

        if login_screen:
            try:
                login_screen.close()
            except Exception:
                logger.debug(
                    "[ACTION_%s] Nao foi possivel fechar sessao SAP",
                    action.id,
                    exc_info=True,
                )


class ActionQueueProcessor:
    """Processa a fila SAP usando o backend como camada de persistencia."""

    def __init__(
        self,
        queue_client: BackendApiClient | None = None,
        sap_runner: SapIw52ActionRunner | None = None,
    ) -> None:
        self.queue_client = queue_client or BackendApiClient()
        self.sap_runner = sap_runner or SapIw52ActionRunner()
        self.results: list[ExecutionResult] = []

    def buscar_acoes_pendentes(self) -> list[QueuedAction]:
        return [QueuedAction.from_api(action) for action in self.queue_client.list_pending_actions()]

    def processar_acao(self, action: QueuedAction) -> ExecutionResult:
        try:
            self.validar_acao(action)
            self.sap_runner.executar_acao_da_fila(action)
            return ExecutionResult(action_id=action.id, success=True)
        except Exception as exc:
            error_message = f"Erro durante execucao: {exc}"
            logger.error("[ACTION_%s] %s", getattr(action, "id", "?"), error_message)
            return ExecutionResult(
                action_id=getattr(action, "id", 0),
                success=False,
                error_message=error_message,
            )

    def registrar_resultado(self, action: QueuedAction, result: ExecutionResult) -> ExecutionResult:
        try:
            updated_action = self.queue_client.complete_action_result(
                action.id,
                success=result.success,
                error_message=result.error_message,
            )
            final_status = updated_action.get("status")
            final_error = updated_action.get("error_message")
            logger.info("[ACTION_%s] Resultado persistido via backend com status %s", action.id, final_status)

            if final_status == "executado":
                return result
            return ExecutionResult(action.id, False, final_error or result.error_message)

        except Exception as exc:
            error_message = f"Erro ao persistir resultado da acao: {exc}"
            logger.error("[ACTION_%s] %s", action.id, error_message)
            self._try_mark_error(action.id, error_message)
            return ExecutionResult(action.id, False, error_message)

    def process_queue(self) -> dict[str, Any]:
        self.results = []
        pending_actions = self.buscar_acoes_pendentes()

        logger.info("=== INICIANDO PROCESSAMENTO DE FILA ===")
        logger.info("Encontradas %s acoes pendentes", len(pending_actions))

        for action in pending_actions:
            logger.info("--- Processando acao %s (%s) ---", action.id, action.action)
            result = self.processar_acao(action)
            result = self.registrar_resultado(action, result)
            self.results.append(result)
            logger.info("--- Acao %s finalizada ---", action.id)

        report = self._build_report(
            total=len(pending_actions),
            success=sum(1 for result in self.results if result.success),
            failed=sum(1 for result in self.results if not result.success),
        )
        logger.info("=== FIM DO PROCESSAMENTO ===")
        return report

    def validar_acao(self, action: QueuedAction) -> None:
        if not action.action:
            raise ActionValidationError("Campo 'action' e obrigatorio")
        if action.action not in {item.value for item in QueueAction}:
            raise ActionValidationError(f"Tipo de acao invalido: {action.action}")
        if not action.registration:
            raise ActionValidationError("Campo 'registration' e obrigatorio")
        if action.region not in REGION_SAP_CONFIG:
            raise ActionValidationError(f"Regiao invalida ou sem configuracao SAP: {action.region}")
        if not action.note:
            raise ActionValidationError("Campo 'note' e obrigatorio para executar acao no SAP")

        if action.action in {QueueAction.PROCEDER.value, QueueAction.IMPROCEDER.value}:
            require_text(action.texto_resultado, "texto_resultado", action)
            require_text(action.texto_medida, "texto_medida", action)

    def _try_mark_error(self, action_id: int, error_message: str) -> None:
        try:
            self.queue_client.complete_action_result(action_id, success=False, error_message=error_message)
        except Exception:
            logger.error("[ACTION_%s] Nao foi possivel registrar erro na fila", action_id, exc_info=True)

    def _build_report(self, total: int, success: int, failed: int) -> dict[str, Any]:
        return {
            "timestamp": datetime.now().isoformat(),
            "total_processed": total,
            "success_count": success,
            "failed_count": failed,
            "results": [result.to_dict() for result in self.results],
            "message": f"Processadas {total} acoes: {success} sucesso, {failed} erro",
        }


def run_queue_processor() -> dict[str, Any]:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    )
    return ActionQueueProcessor().process_queue()

if __name__ == "__main__":
    report = run_queue_processor()
    print(report)
