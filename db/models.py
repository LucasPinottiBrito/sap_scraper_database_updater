from sqlalchemy import Column, Integer, String, ForeignKey
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()

class Note(Base):
    __tablename__ = "sapAutoTPendingNotes"

    id = Column(Integer, primary_key=True, autoincrement=True)
    note_number = Column(String(255), nullable=True)
    created_at = Column(String(255), nullable=True)
    date = Column(String(255), nullable=True)
    priority_text = Column(String(255), nullable=True)
    group_ = Column("group", String(255), nullable=True)
    code_text = Column(String(255), nullable=True)
    code_group_text = Column(String(255), nullable=True)
    city = Column(String(255), nullable=True)
    description = Column(String(1024), nullable=True)
    description_detail = Column(String(4096), nullable=True)
    business_partner_id = Column(String(255), nullable=True)
    inst = Column(String(255), nullable=True)
    email = Column(String(255), nullable=True)
    sms = Column(String(255), nullable=True)
    cod_contact = Column(String(2), nullable=True)

    def to_dict(self):
        return {
            "id": self.id,
            "note_number": self.note_number,
            "created_at": self.created_at,
            "date": self.date,
            "priority_text": self.priority_text,
            "group": self.group_,
            "code_text": self.code_text,
            "code_group_text": self.code_group_text,
            "city": self.city,
            "description": self.description,
            "description_detail": self.description_detail,
            "business_partner_id": self.business_partner_id,
            "inst": self.inst,
        }


class Attachment(Base):
    __tablename__ = "sapAutoTAttachments"

    id = Column(Integer, primary_key=True, autoincrement=True)
    note_id = Column(Integer, ForeignKey("sapAutoTPendingNotes.id", ondelete="CASCADE"), nullable=False)
    url = Column(String(1024), nullable=False)
    created_at = Column(String(255), nullable=True)


    def to_dict(self):
        return {
            "id": self.id,
            "note_id": self.note_id,
            "url": self.url,
            "created_at": self.created_at,
        }
