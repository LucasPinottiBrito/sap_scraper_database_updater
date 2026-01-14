from typing import Optional, List, Dict, Any
from sqlalchemy.orm import Session
from db.config import get_session
from db.models import Note, Attachment

class AttachmentRepository:
    def __init__(self, session: Optional[Session] = None):
        self.session = session or get_session()
        self._close = session is None

    def create(self, note_id: int, url: str, created_at: Optional[str] = None) -> Attachment:
        note = self.session.query(Note).filter_by(id=note_id).one_or_none()
        if not note:
            raise ValueError("Note not found")
        attach = Attachment(note_id=note_id, url=url, created_at=created_at)
        self.session.add(attach)
        self.session.commit()
        self.session.refresh(attach)
        return attach

    def get(self, attach_id: int) -> Optional[Attachment]:
        return self.session.query(Attachment).filter_by(id=attach_id).one_or_none()

    def list_by_note(self, note_id: int) -> List[Attachment]:
        return self.session.query(Attachment).filter_by(note_id=note_id).all()

    def update(self, attach_id: int, updates: Dict[str, Any]) -> Attachment:
        a = self.get(attach_id)
        if not a:
            raise ValueError("Attachment not found")
        for k, v in updates.items():
            setattr(a, k, v)
        self.session.commit()
        self.session.refresh(a)
        return a

    def delete(self, attach_id: int) -> None:
        a = self.get(attach_id)
        if not a:
            raise ValueError("Attachment not found")
        self.session.delete(a)
        self.session.commit()

    def __del__(self):
        if self._close:
            self.session.close()
