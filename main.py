from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import time

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


if __name__ == "__main__":
    print("SAP Module")
    while ambiente not in [1, 2, 3, 4]:
        ambiente = int(input("Select environment (1- EP1, 2- EP2, 3- CP1-500, 4- CP1-600): "))
    while not usuario:
        usuario = input("Enter your username: ")
    while not senha:
        senha = input("Enter your password: ")

    ambiente_selecionado = ambiente_map[ambiente-1]['name']
    regiao_selecionada = ambiente_map[ambiente-1]['region']

    print(f"Environment: {ambiente_selecionado}, User: {usuario}, Password: {'*' * len(senha)}")

    logon = SapLogonScreen()
    login = logon.loadSystem(regiao_selecionada, ambiente_selecionado)
    home = login.login(usuario, senha, regiao_selecionada, ambiente_selecionado)
    iw67Screen = home.openTransaction("iw67")
    notes_screen = iw67Screen.openNotesScreen(regiao_selecionada)
    notes = notes_screen.getNotes()
    iw67Screen.back()
    iw52Screen = home.openTransaction("iw52")
    for note_number in notes:
        print(f"Processing Note Number: {note_number.get('nota')}")
        iw52NoteScreen = iw52Screen.openNote(note_number.get("nota"))
        attachments = iw52NoteScreen.get_attachments(download_files=False)
        note_number['attachments'] = attachments
        print(f"Note Number: {note_number}")
        iw52NoteScreen.back()
        # break  # Remove this break to process all notes
    
    df = pd.DataFrame(notes)
    df.to_csv("notes_with_attachments.csv", index=False, encoding='utf-8-sig')

    iw52Screen.close()