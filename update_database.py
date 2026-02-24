from db.repositories.TablesUpdatedAt import TablesUpdatedAtRepository
from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import datetime
from db.repositories.Note import NoteRepository
from db.repositories.BIEntity import BIEntityRepository
from db.repositories.Attachment import AttachmentRepository
from db.repositories.UpdateJob import UpdateJobRepository
from dotenv import load_dotenv
import os

load_dotenv()

config_file_path = "config.txt"
usuario = os.getenv("SAP_USER", "")
senha = os.getenv("SAP_PASSWORD_EP1", "")
senha_ep2 = os.getenv("SAP_PASSWORD_EP2", "")
ambiente = ""

updated_at_repo = TablesUpdatedAtRepository()


def get_notes_from_environment(ambiente_selecionado: str, regiao_selecionada: str, senha) -> pd.DataFrame:
    logon = SapLogonScreen()
    login = logon.loadSystem(regiao_selecionada, ambiente_selecionado)
    home = login.login(usuario, senha, regiao_selecionada, ambiente_selecionado)
    iw67Screen = home.openTransaction("iw67")
    try:
        notes_screen = iw67Screen.openNotesScreen(regiao_selecionada, noteType="REC")
        rec_notes = notes_screen.getNotes("REC")
    except Exception as e:
        print(f"No notes found for REC type: {e}")
    
    try:
        notes_screen = iw67Screen.openNotesScreen(regiao_selecionada, noteType="PRC")
        prc_notes = notes_screen.getNotes("PRC")
    except Exception as e:
        prc_notes = []
        print(f"No notes found for PRC type: {e}")
    
    try:
        notes_screen = iw67Screen.openNotesScreen(regiao_selecionada, noteType="SC/RC")
        sc_rc_notes = notes_screen.getNotes("SC/RC")
    except Exception as e:
        print(f"No notes found for SC/RC type: {e}")
    
    try:
        notes_screen = iw67Screen.openNotesScreen(regiao_selecionada, noteType="OVD")
        ovd_notes = notes_screen.getNotes("OVD")
    except Exception as e:
        print(f"No notes found for OVD type: {e}")

    notes = rec_notes + prc_notes + sc_rc_notes + ovd_notes
    df = pd.DataFrame(notes)
    notes_screen.close()

    return df

def get_notes_details_and_attachments(ambiente_selecionado: str, regiao_selecionada: str, notes: list, senha: str) -> list:
    logon = SapLogonScreen()
    login = logon.loadSystem(regiao_selecionada, ambiente_selecionado)
    home = login.login(usuario, senha, regiao_selecionada, ambiente_selecionado)
    iw52Screen = home.openTransaction("iw52")
    for note_number in notes:
        print(f"Processing Note Number: {note_number.get('note_number')}")
        
        try:
            iw52NoteScreen = iw52Screen.openNote(note_number.get("note_number"))
            note_model = "CI" if note_number.get("note_type") == "Recurso" else "NA"
            details = iw52NoteScreen.getNoteDetails(note_model=note_model)
            note_number['description_detail'] = details.get("description_text")
            note_number['email'] = details.get("contato_email")
            note_number['sms'] = details.get("contato_sms")
            note_number['cod_contact'] = details.get("cod_contato")
            note_number['inst'] = details.get("inst")
        except Exception as e:
            print(f"Error opening note {note_number.get('note_number')}: {e}")
            continue
        
        attachments = iw52NoteScreen.get_attachments(download_files=False, folder_path_to_download=f"./attachments/{ambiente_selecionado}/{note_number.get('note_number')}/")
        note_number['attachments'] = attachments
        iw52NoteScreen.back()

    df = pd.DataFrame(notes)
    iw52Screen.close()

    return df

def save_note_details_and_attachments(ambiente_selecionado: str, regiao_selecionada: str, notes: list, senha: str):
    logon = SapLogonScreen()
    login = logon.loadSystem(regiao_selecionada, ambiente_selecionado)
    home = login.login(usuario, senha, regiao_selecionada, ambiente_selecionado)
    iw52Screen = home.openTransaction("iw52")
    for note_number in notes:
        note_repo = NoteRepository()
        nota = note_repo.get_from_number(note_number.get("note_number"))
        if nota:
            print(f"Note {note_number.get('note_number')} already exists in the database. Skipping.")
            continue

        print(f"Processing Note Number: {note_number.get('note_number')}")
        
        try:
            iw52NoteScreen = iw52Screen.openNote(note_number.get("note_number"))
            note_model = "CI" if note_number.get("note_type") == "Recurso" else "NA"
            details = iw52NoteScreen.getNoteDetails(note_model=note_model)
            nota = note_repo.create(
                {
                    "note_number": note_number.get("note_number"),
                    "created_at": note_number.get("created_at"),
                    "conclusion_date": note_number.get("conclusion_date"),
                    "priority_text": note_number.get("priority_text"),
                    "group": note_number.get("note_type"),
                    "code_text": note_number.get("code_text"),
                    "city": note_number.get("city"),
                    "description": note_number.get("description"),
                    "description_detail": details.get("description_text"),
                    "business_partner_id": note_number.get("business_partner_id"),
                    "email": details.get("contato_email"),
                    "sms": details.get("contato_sms"),
                    "cod_contact": details.get("cod_contato"),
                    "inst": details.get("inst")
                }
            )
            updated_at_repo.update_table_timestamp("sapAutoTPendingNotes")

        except Exception as e:
            print(f"Error opening note {note_number.get('note_number')}: {e}")
            continue
        
        attachments = iw52NoteScreen.get_attachments(download_files=False, folder_path_to_download=f"./attachments/{ambiente_selecionado}/{note_number.get('note_number')}/")
        if attachments:
            attach_repo = AttachmentRepository()
            for attach in attachments:
                try:
                    attach_repo.create(
                        note_id=nota.id,
                        url=attach,
                        created_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    )
                    updated_at_repo.update_table_timestamp("sapAutoTAttachments")
                    print(f"Attachment {attach} added to database.")
                except Exception as e:
                    print(f"Error adding attachment {attach} to database: {e}")
        
        note_number['attachments'] = attachments
        iw52NoteScreen.back()

    iw52Screen.close()


def run_database_update():
    print("SP")
    sp_notes = get_notes_from_environment("EP1", "SP", senha)
    sp_notes = get_notes_details_and_attachments("EP1", "SP", sp_notes.to_dict(orient='records'), senha)
    sp_notes["region"] = "SP"
    
    print("ES")
    es_notes = get_notes_from_environment("EP2", "ES", senha_ep2)
    es_notes = get_notes_details_and_attachments("EP2", "ES", es_notes.to_dict(orient='records'), senha_ep2)
    es_notes["region"] = "ES"

    all_notes = pd.concat([sp_notes, es_notes], ignore_index=True)
    save_note_details_and_attachments("EP1", "SP", sp_notes.to_dict(orient='records'), senha)
    save_note_details_and_attachments("EP2", "ES", es_notes.to_dict(orient='records'), senha_ep2)

    notes_repo = NoteRepository()
    notes_repo.delete_all(notes_to_keep=all_notes.get("note_number").tolist())

    print("Process completed.")