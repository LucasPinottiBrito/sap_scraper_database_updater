from typing import Optional, List, Dict, Any
from sqlalchemy.orm import Session
from db.config import get_session
from db.models import Note

class NoteRepository:
    def __init__(self, session: Optional[Session] = None):
        self.session = session or get_session()
        self._close = session is None

    def create(self, payload: Dict[str, Any]) -> Note:
        p = dict(payload)
        if "group" in p:
            p["group_"] = p.pop("group")
        note = Note(**p)
        self.session.add(note)
        self.session.commit()
        self.session.refresh(note)
        return note

    def get(self, note_id: int) -> Optional[Note]:
        return self.session.query(Note).filter_by(id=note_id).one_or_none()

    def get_from_number(self, note_number: str) -> Optional[Note]:
        return self.session.query(Note).filter_by(note_number=note_number).one_or_none()

    def list(self, filters: Optional[Dict[str, Any]] = None,
             limit: Optional[int] = None,
             offset: Optional[int] = None) -> List[Note]:
        q = self.session.query(Note)
        if filters:
            f = dict(filters)
            if "group" in f:
                f["group_"] = f.pop("group")
            q = q.filter_by(**f)
        if offset:
            q = q.offset(offset)
        if limit:
            q = q.limit(limit)
        return q.all()

    def update(self, note_id: int, updates: Dict[str, Any]) -> Note:
        note = self.get(note_id)
        if not note:
            raise ValueError("Note not found")
        u = dict(updates)
        if "group" in u:
            u["group_"] = u.pop("group")
        for k, v in u.items():
            setattr(note, k, v)
        self.session.commit()
        self.session.refresh(note)
        return note

    def delete(self, note_id: int) -> None:
        note = self.get(note_id)
        if not note:
            raise ValueError("Note not found")
        self.session.delete(note)
        self.session.commit()

    def delete_all(self, notes_to_keep: List[str] = None) -> None:
        if notes_to_keep:
            self.session.query(Note).filter(~Note.note_number.in_(notes_to_keep)).delete(synchronize_session=False)
        else:
            self.session.query(Note).delete()
        self.session.commit()

    def __del__(self):
        if self._close:
            self.session.close()
