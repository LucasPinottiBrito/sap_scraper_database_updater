from db.repositories.TablesUpdatedAt import TablesUpdatedAtRepository
from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
from db.repositories.BIEntity import BIEntityRepository
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

def run_bi_update():
    print("SP")
    sp_notes = get_notes_from_environment("EP1", "SP", senha)
    sp_notes["region"] = "SP"
    
    print("ES")
    es_notes = get_notes_from_environment("EP2", "ES", senha_ep2)
    es_notes["region"] = "ES"

    all_notes = pd.concat([sp_notes, es_notes], ignore_index=True)
    all_notes.rename(columns={"note_type": "TipoNota"}, inplace=True)
    all_notes.rename(columns={"note_number": "NotaSAP"}, inplace=True)
    all_notes.rename(columns={"created_at": "DataCriacao"}, inplace=True)
    all_notes.rename(columns={"conclusion_date": "ConclusaoDesejada"} , inplace=True)
    all_notes.rename(columns={"region": "Estado"} , inplace=True)

    replace_notes = all_notes.to_dict(orient='records')
    bi_entity_repo = BIEntityRepository()
    bi_entity_repo.replace_all(replace_notes)
    updated_at_repo.update_table_timestamp("NotasAbertasTable")

    print("Process completed.")