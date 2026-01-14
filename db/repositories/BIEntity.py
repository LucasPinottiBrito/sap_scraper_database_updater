from typing import List, Dict, Any, Optional
from sqlalchemy.orm import Session
from sqlalchemy import delete, insert
from db.config import get_session
from db.tables import NotasAbertasTable


class BIEntityRepository:
    def __init__(self, session: Optional[Session] = None):
        self.session = session or get_session()
        self._close = session is None

    def replace_all(self, entities: List[Dict[str, Any]]) -> None:

        try:
            # DELETE total
            self.session.execute(delete(NotasAbertasTable))

            # BULK INSERT
            if entities:
                self.session.execute(
                    insert(NotasAbertasTable),
                    entities
                )

            self.session.commit()

        except Exception:
            self.session.rollback()
            raise

        finally:
            if self._close:
                self.session.close()
