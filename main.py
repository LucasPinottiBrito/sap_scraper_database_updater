from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import datetime
from db.models import Base
from db.config import engine
from db.note_repo import NoteRepository
from db.bi_entity_repo import BIEntityRepository
from db.attach_repo import AttachmentRepository

Base.metadata.create_all(engine)

config_file_path = "config.txt"
usuario = ""
senha = ""
ambiente = ""

try:
    config = open(config_file_path, mode='rb')
    usuario = config.readline().decode('utf-8').strip() or ""
    senha = config.readline().decode('utf-8').strip() or ""
except FileNotFoundError:
    print(f"Configuration file not found at {config_file_path}. Please create the file with the required parameters.")


ambiente_map = [
    {"region": "SP", "name": "EP1"},
    {"region": "ES", "name": "EP2"},
    {"region": "SP", "name": "CP1"},
    {"region": "ES", "name": "CP1"},
]

note_repo = NoteRepository()

def update_notes_from_environment(ambiente_selecionado: str, regiao_selecionada) -> pd.DataFrame:

    print(f"Environment: {ambiente_selecionado}, User: {usuario}, Password: {'*' * len(senha)}")

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
    print(f"Total notes fetched: {len(notes)}")

    df = pd.DataFrame(notes)
    
    return df
    # iw52Screen = home.openTransaction("iw52")
    # for note_number in notes:
    #     nota = note_repo.get_from_number(note_number.get("note_number"))
    #     if nota:
    #         print(f"Note {note_number.get('note_number')} already exists in the database. Skipping.")
    #         continue

    #     print(f"Processing Note Number: {note_number.get('note_number')}")
        
    #     try:
    #         iw52NoteScreen = iw52Screen.openNote(note_number.get("note_number"))
    #         details = iw52NoteScreen.getNoteDetails()
    #         nota = note_repo.create(
    #             {
    #                 "note_number": note_number.get("note_number"),
    #                 "created_at": note_number.get("created_at"),
    #                 "date": note_number.get("date"),
    #                 "priority_text": note_number.get("priority_text"),
    #                 "group": note_number.get("group"),
    #                 "code_text": note_number.get("code_text"),
    #                 "code_group_text": note_number.get("code_group_text"),
    #                 "city": note_number.get("city"),
    #                 "description": note_number.get("description"),
    #                 "description_detail": details.get("description_text"),
    #                 "business_partner_id": note_number.get("business_partner_id"),
    #                 "email": details.get("contato_email"),
    #                 "sms": details.get("contato_sms"),
    #                 "cod_contact": details.get("cod_contato"),
    #                 "inst": details.get("inst")
    #             }
    #         )

    #     except Exception as e:
    #         print(f"Error opening note {note_number.get('note_number')}: {e}")
    #         continue

    #     attachments = iw52NoteScreen.get_attachments(download_files=False, folder_path_to_download=f"./attachments/{ambiente_selecionado}/{note_number.get('note_number')}/")
    #     if attachments:
    #         attach_repo = AttachmentRepository()
    #         for attach in attachments:
    #             try:
    #                 attach_repo.create(
    #                     note_id=nota.id,
    #                     url=attach,
    #                     created_at=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    #                 )
    #                 print(f"Attachment {attach} added to database.")
    #             except Exception as e:
    #                 print(f"Error adding attachment {attach} to database: {e}")
        
    #     note_number['attachments'] = attachments
    #     iw52NoteScreen.back()
    
    # df = pd.DataFrame(notes)
    # today_date = datetime.datetime.now().strftime("%Y%m%d")
    # df.to_csv(f"notes_with_attachments_{today_date}_{ambiente_selecionado}.csv", index=False, encoding='utf-8-sig')

    # iw52Screen.close()

if __name__ == "__main__":
    print("SP")
    sp_notes = update_notes_from_environment("EP1", "SP")
    sp_notes["region"] = "SP"
    print("ES")
    es_notes = update_notes_from_environment("EP2", "ES")
    es_notes["region"] = "ES"

    all_notes = pd.concat([sp_notes, es_notes], ignore_index=True)
    all_notes.rename(columns={"note_type": "TipoNota"}, inplace=True)
    all_notes.rename(columns={"note_number": "NotaSAP"}, inplace=True)
    all_notes.rename(columns={"created_at": "DataCriacao"}, inplace=True)
    all_notes.rename(columns={"conclusion_date": "ConclusaoDesejada"} , inplace=True)
    all_notes.rename(columns={"region": "Estado"} , inplace=True)

    replace_notes = all_notes.to_dict(orient='records')
    print(f"Total notes to replace in DB: {len(replace_notes)}")
    print(replace_notes)
    bi_entity_repo = BIEntityRepository()
    bi_entity_repo.replace_all(replace_notes)