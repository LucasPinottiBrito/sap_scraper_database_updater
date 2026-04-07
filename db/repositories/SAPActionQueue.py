from datetime import datetime
from typing import Optional, List, Dict, Any
from sqlalchemy.orm import Session
from db.config import get_session
from db.models import SAPActionQueue


class SAPActionQueueRepository:
    
    def __init__(self, session: Optional[Session] = None):
        self.session = session or get_session()
        self._close = session is None

    def create(self, payload: Dict[str, Any]) -> SAPActionQueue:
        """Criar nova ação na fila"""
        action = SAPActionQueue(**payload)
        self.session.add(action)
        self.session.commit()
        self.session.refresh(action)
        return action

    def get(self, action_id: int) -> Optional[SAPActionQueue]:
        """Buscar ação por ID"""
        return self.session.query(SAPActionQueue).filter_by(id=action_id).one_or_none()

    def list_pending(self) -> List[SAPActionQueue]:
        """Buscar todas ações com status='pendente'"""
        return self.session.query(SAPActionQueue).filter_by(status="pendente").all()

    def list_all(self, filters: Optional[Dict[str, Any]] = None) -> List[SAPActionQueue]:
        """Listar ações com filtros opcionais"""
        q = self.session.query(SAPActionQueue)
        if filters:
            q = q.filter_by(**filters)
        return q.all()

    def get_by_note(self, note_number: str) -> Optional[SAPActionQueue]:
        """Buscar ação por número da nota (note field)"""
        return self.session.query(SAPActionQueue).filter_by(note=note_number).one_or_none()

    def update_status(
        self, 
        action_id: int, 
        status: str, 
        error_message: Optional[str] = None
    ) -> SAPActionQueue:
        """
        Atualizar status de uma ação
        
        Args:
            action_id: ID da ação
            status: novo status ('ok', 'erro', 'pendente')
            error_message: mensagem de erro (opcional)
        
        Returns:
            Ação atualizada
        """
        action = self.get(action_id)
        if not action:
            raise ValueError(f"Ação {action_id} não encontrada")
        
        action.status = status
        action.conclusion_date = datetime.now()
        
        if error_message:
            action.error_message = error_message
        
        self.session.commit()
        self.session.refresh(action)
        return action

    def get_by_registration_and_action(
        self, 
        registration: str, 
        action_type: str
    ) -> Optional[SAPActionQueue]:
        """Buscar ação por matrícula e tipo de ação"""
        return self.session.query(SAPActionQueue).filter_by(
            registration=registration,
            action=action_type
        ).one_or_none()

    def __del__(self):
        if self._close:
            self.session.close()
