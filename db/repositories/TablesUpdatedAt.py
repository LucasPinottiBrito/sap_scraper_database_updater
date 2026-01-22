from typing import Optional, List, Dict, Any
from sqlalchemy import func
from sqlalchemy.orm import Session
from db.config import get_session
from db.models import Note, TablesUpdatedAt, TablesUpdatedAt

class TablesUpdatedAtRepository:
    def __init__(self, session: Optional[Session] = None):
        self.session = session or get_session()
        self._close = session is None

    def get(self, table_name: str) -> Optional[TablesUpdatedAt]:
        return self.session.query(TablesUpdatedAt).filter_by(table_name=table_name).one_or_none()

    def create(self, payload: Dict[str, Any]) -> TablesUpdatedAt:
        p = dict(payload)
        if "group" in p:
            p["group_"] = p.pop("group")
        tables_updated_at = TablesUpdatedAt(**p)
        self.session.add(tables_updated_at)
        self.session.commit()
        self.session.refresh(tables_updated_at)
        return tables_updated_at

    def list(self, filters: Optional[Dict[str, Any]] = None,
             limit: Optional[int] = None,
             offset: Optional[int] = None) -> List[TablesUpdatedAt]:
        q = self.session.query(TablesUpdatedAt)
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

    def update(self, table_name: str, updates: Dict[str, Any]) -> TablesUpdatedAt:
        tables_updated_at = self.get(table_name)
        if not tables_updated_at:
            raise ValueError("TablesUpdatedAt not found")
        u = dict(updates)
        if "group" in u:
            u["group_"] = u.pop("group")
        for k, v in u.items():
            setattr(tables_updated_at, k, v)
        self.session.commit()
        self.session.refresh(tables_updated_at)
        return tables_updated_at

    def update_table_timestamp(self, table_name: str) -> TablesUpdatedAt:
        tables_updated_at = self.get(table_name)
        if not tables_updated_at:
            tables_updated_at = self.create({
                "table_name": table_name,
                "updated_at": func.now()
            })
        else:
            tables_updated_at.updated_at = func.now()
            self.session.commit()
            self.session.refresh(tables_updated_at)
        return tables_updated_at

    def delete(self, table_name: str) -> None:
        tables_updated_at = self.get(table_name)
        if not tables_updated_at:
            raise ValueError("TablesUpdatedAt not found")
        self.session.delete(tables_updated_at)
        self.session.commit()

    def __del__(self):
        if self._close:
            self.session.close()
