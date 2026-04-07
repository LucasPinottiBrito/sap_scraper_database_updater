"""
Action Queue Processor

Responsável por:
1. Buscar ações pendentes de sapAutoTActionQueue
2. Validar referências de notas
3. Mapear ações para funções específicas
4. Executar ações e registrar resultado
5. Deletar notas quando ação é finalizada com sucesso

Ações suportadas:
- proceder: Avançar com o processamento da nota
- improceder: Rejeitar o processamento
- transferir_semci: Transferir para SEMCI
- transferir_mmgd: Transferir para MMGD
- encerrar: Encerrar o processamento
"""

from datetime import datetime
from typing import Dict, Any, Optional, Callable, List
import logging
from db.repositories.SAPActionQueue import SAPActionQueueRepository
from db.repositories.Note import NoteRepository
from db.models import SAPActionQueue

logger = logging.getLogger(__name__)


class ExecutionResult:
    """Resultado de execução de uma ação"""
    def __init__(self, action_id: int, success: bool, error_message: Optional[str] = None):
        self.action_id = action_id
        self.success = success
        self.error_message = error_message
        self.timestamp = datetime.now()

    def to_dict(self):
        return {
            "action_id": self.action_id,
            "success": self.success,
            "error_message": self.error_message,
            "timestamp": self.timestamp.isoformat()
        }


class ActionExecutor:
    """
    Executor de ações
    
    Define e executa as 5 ações suportadas. Cada ação recebe um contexto
    com informações da nota e ação.
    """
    
    @staticmethod
    def proceder(context: Dict[str, Any]) -> bool:
        """
        Ação: PROCEDER
        Avança com o processamento da nota.
        Pode envolver: validação, aprovação, marcar como processada.
        
        Args:
            context: {
                'action_id': int,
                'registration': str (matrícula/cpf),
                'note_number': str,
                'region': str,
                'note_text': str,
                'note_id': int
            }
        
        Returns:
            bool: True se sucesso, False se falha
        """
        try:
            logger.info(f"[PROCEDER] Iniciando para nota {context.get('note_number')}")
            # TODO: Implementar lógica específica de processamento
            # Exemplo: chamar API SAP, atualizar status, enviar notificação
            logger.info(f"[PROCEDER] Concluído com sucesso para nota {context.get('note_number')}")
            return True
        except Exception as e:
            logger.error(f"[PROCEDER] Erro: {str(e)}")
            raise

    @staticmethod
    def improceder(context: Dict[str, Any]) -> bool:
        """
        Ação: IMPROCEDER
        Rejeita o processamento da nota com motivo.
        
        Args:
            context: {dicionário com informações da ação}
        
        Returns:
            bool: True se sucesso, False se falha
        """
        try:
            logger.info(f"[IMPROCEDER] Iniciando para nota {context.get('note_number')}")
            # TODO: Implementar lógica de rejeição
            # Exemplo: registrar motivo, notificar responsável, arquivar nota
            logger.info(f"[IMPROCEDER] Concluído com sucesso para nota {context.get('note_number')}")
            return True
        except Exception as e:
            logger.error(f"[IMPROCEDER] Erro: {str(e)}")
            raise

    @staticmethod
    def transferir_semci(context: Dict[str, Any]) -> bool:
        """
        Ação: TRANSFERIR_SEMCI
        Transfere o processamento para SEMCI.
        
        Args:
            context: {dicionário com informações da ação}
        
        Returns:
            bool: True se sucesso, False se falha
        """
        try:
            logger.info(f"[TRANSFERIR_SEMCI] Iniciando para nota {context.get('note_number')}")
            # TODO: Implementar lógica de transferência
            # Exemplo: criar tarefa em SEMCI, notificar equipe, atualizar status
            logger.info(f"[TRANSFERIR_SEMCI] Concluído com sucesso para nota {context.get('note_number')}")
            return True
        except Exception as e:
            logger.error(f"[TRANSFERIR_SEMCI] Erro: {str(e)}")
            raise

    @staticmethod
    def transferir_mmgd(context: Dict[str, Any]) -> bool:
        """
        Ação: TRANSFERIR_MMGD
        Transfere o processamento para MMGD.
        
        Args:
            context: {dicionário com informações da ação}
        
        Returns:
            bool: True se sucesso, False se falha
        """
        try:
            logger.info(f"[TRANSFERIR_MMGD] Iniciando para nota {context.get('note_number')}")
            # TODO: Implementar lógica de transferência
            # Exemplo: criar tarefa em MMGD, notificar equipe, atualizar status
            logger.info(f"[TRANSFERIR_MMGD] Concluído com sucesso para nota {context.get('note_number')}")
            return True
        except Exception as e:
            logger.error(f"[TRANSFERIR_MMGD] Erro: {str(e)}")
            raise

    @staticmethod
    def encerrar(context: Dict[str, Any]) -> bool:
        """
        Ação: ENCERRAR
        Encerra definitivamente o processamento da nota.
        
        Args:
            context: {dicionário com informações da ação}
        
        Returns:
            bool: True se sucesso, False se falha
        """
        try:
            logger.info(f"[ENCERRAR] Iniciando para nota {context.get('note_number')}")
            # TODO: Implementar lógica de encerramento
            # Exemplo: marcar como encerrada, arquivar, gerar relatório
            logger.info(f"[ENCERRAR] Concluído com sucesso para nota {context.get('note_number')}")
            return True
        except Exception as e:
            logger.error(f"[ENCERRAR] Erro: {str(e)}")
            raise


class ActionQueueProcessor:
    """
    Processador de fila de ações
    
    Workflow:
    1. Busca ações com status='pendente'
    2. Valida referências e dados
    3. Executa função específica baseada no tipo de ação
    4. Atualiza status: 'ok' (sucesso) ou 'erro' (falha)
    5. Se status='ok', deleta a nota associada
    6. Continua com próxima ação mesmo se uma falhar
    """
    
    # Mapeamento de tipos de ação para funções executoras
    ACTION_HANDLERS: Dict[str, Callable] = {
        "proceder": ActionExecutor.proceder,
        "improceder": ActionExecutor.improceder,
        "transferir_semci": ActionExecutor.transferir_semci,
        "transferir_mmgd": ActionExecutor.transferir_mmgd,
        "encerrar": ActionExecutor.encerrar,
    }
    
    def __init__(self):
        self.action_repo = SAPActionQueueRepository()
        self.note_repo = NoteRepository()
        self.results: List[ExecutionResult] = []

    def fetch_pending_actions(self) -> List[SAPActionQueue]:
        """Busca todas as ações com status='pendente'"""
        return self.action_repo.list_pending()

    def validate_action(self, action: SAPActionQueue) -> tuple[bool, Optional[str]]:
        """
        Valida se uma ação pode ser executada
        
        Validações:
        - Tipo de ação é válido?
        - Nota associada existe?
        - Dados obrigatórios presentes?
        
        Returns:
            (valid, error_message)
        """
        # Validar tipo de ação
        if action.action not in self.ACTION_HANDLERS:
            return False, f"Tipo de ação inválido: {action.action}"
        
        # Validar dados obrigatórios
        if not action.registration:
            return False, "Campo 'registration' é obrigatório"
        
        if not action.region:
            return False, "Campo 'region' é obrigatório"
        
        return True, None

    def build_execution_context(self, action: SAPActionQueue) -> Dict[str, Any]:
        """
        Constrói contexto para executar a ação
        
        Inclui informações da ação e dados da nota relacionada
        """
        context = {
            "action_id": action.id,
            "registration": action.registration,
            "action_type": action.action,
            "note_text": action.note,
            "region": action.region,
            "created_at": action.created_at,
        }
        
        # Se houver identificação da nota, buscar dados adicionais
        if action.note:
            note = self.note_repo.get_from_number(action.note)
            if note:
                context["note_id"] = note.id
                context["note_number"] = note.note_number
                context["note_data"] = note.to_dict()
        
        return context

    def execute_action(self, action: SAPActionQueue) -> ExecutionResult:
        """
        Executa uma ação específica
        
        Args:
            action: Ação da fila
        
        Returns:
            ExecutionResult com resultado da execução
        """
        try:
            # Validar ação
            is_valid, error_msg = self.validate_action(action)
            if not is_valid:
                logger.error(f"[ACTION_{action.id}] Validação falhou: {error_msg}")
                return ExecutionResult(action.id, False, error_msg)
            
            # Construir contexto
            context = self.build_execution_context(action)
            
            # Buscar executor
            handler = self.ACTION_HANDLERS.get(action.action)
            
            logger.info(f"[ACTION_{action.id}] Iniciando: {action.action} para {action.registration}")
            
            # Executar handler
            success = handler(context)
            
            # Criar resultado
            result = ExecutionResult(action.id, success)
            
            logger.info(f"[ACTION_{action.id}] Concluído com sucesso")
            
            return result
            
        except Exception as e:
            error_message = f"Erro durante execução: {str(e)}"
            logger.error(f"[ACTION_{action.id}] {error_message}")
            return ExecutionResult(action.id, False, error_message)

    def update_action_status(
        self, 
        action_id: int, 
        success: bool, 
        error_message: Optional[str] = None
    ) -> None:
        """
        Atualiza status de uma ação no banco
        
        Args:
            action_id: ID da ação
            success: True = 'ok', False = 'erro'
            error_message: Mensagem de erro (se houver)
        """
        status = "ok" if success else "erro"
        try:
            self.action_repo.update_status(action_id, status, error_message)
            logger.info(f"[ACTION_{action_id}] Status atualizado para: {status}")
        except Exception as e:
            logger.error(f"[ACTION_{action_id}] Erro ao atualizar status: {str(e)}")

    def delete_related_note(self, action: SAPActionQueue) -> bool:
        """
        Deleta a nota associada à ação (apenas após sucesso)
        
        Args:
            action: Ação que foi executada com sucesso
        
        Returns:
            bool: True se deletada, False se erro
        """
        try:
            if not action.note:
                logger.warning(f"[ACTION_{action.id}] Nenhuma nota para deletar")
                return False
            
            # Buscar nota por número
            note = self.note_repo.get_from_number(action.note)
            if not note:
                logger.warning(f"[ACTION_{action.id}] Nota {action.note} não encontrada")
                return False
            
            # Deletar
            self.note_repo.delete(note.id)
            logger.info(f"[ACTION_{action.id}] Nota {action.note} deletada com sucesso")
            return True
            
        except Exception as e:
            logger.error(f"[ACTION_{action.id}] Erro ao deletar nota: {str(e)}")
            return False

    def process_queue(self) -> Dict[str, Any]:
        """
        Processa todas as ações pendentes
        
        Workflow:
        1. Busca ações pendentes
        2. Executa cada ação isoladamente
        3. Atualiza status no banco
        4. Se sucesso, deleta nota
        5. Continua mesmo se uma falhar
        
        Returns:
            Relatório de execução com estatísticas
        """
        self.results = []
        
        logger.info("=== INICIANDO PROCESSAMENTO DE FILA ===")
        
        # Buscar ações
        pending_actions = self.fetch_pending_actions()
        logger.info(f"Encontradas {len(pending_actions)} ações pendentes")
        
        if not pending_actions:
            logger.info("Nenhuma ação para processar")
            return self._build_report(0, 0, 0)
        
        # Processar cada ação
        for action in pending_actions:
            logger.info(f"\n--- Processando ação {action.id} ---")
            
            # Executar
            result = self.execute_action(action)
            self.results.append(result)
            
            # Atualizar status
            self.update_action_status(action.id, result.success, result.error_message)
            
            # Se sucesso, deletar nota
            if result.success:
                self.delete_related_note(action)
            
            logger.info(f"--- Ação {action.id} finalizada ---\n")
        
        logger.info("=== FIM DO PROCESSAMENTO ===")
        
        # Construir relatório
        return self._build_report(
            total=len(pending_actions),
            success=sum(1 for r in self.results if r.success),
            failed=sum(1 for r in self.results if not r.success)
        )

    def _build_report(self, total: int, success: int, failed: int) -> Dict[str, Any]:
        """Constrói relatório final"""
        return {
            "timestamp": datetime.now().isoformat(),
            "total_processed": total,
            "success_count": success,
            "failed_count": failed,
            "results": [r.to_dict() for r in self.results],
            "message": f"Processadas {total} ações: {success} sucesso, {failed} erro"
        }

    def __del__(self):
        """Cleanup"""
        try:
            if hasattr(self, 'action_repo'):
                del self.action_repo
            if hasattr(self, 'note_repo'):
                del self.note_repo
        except:
            pass


def run_queue_processor() -> Dict[str, Any]:
    """
    Função de entrada para executar o processador
    
    Pode ser chamada por um scheduler ou manualmente
    
    Returns:
        Relatório de execução
    """
    # Configurar logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )
    
    processor = ActionQueueProcessor()
    return processor.process_queue()
