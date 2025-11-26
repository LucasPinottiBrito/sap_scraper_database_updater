from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import datetime

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
    while ambiente not in [1, 2]:
        ambiente = int(input("Select environment (1- EP1, 2- EP2): "))
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
        try:
            iw52NoteScreen = iw52Screen.openNote(note_number.get("nota"))
        except Exception as e:
            print(f"Error opening note {note_number.get('nota')}: {e}")
            continue
        attachments = iw52NoteScreen.get_attachments(download_files=False, folder_path_to_download=f"./attachments/{ambiente_selecionado}/{note_number.get('nota')}/")
        note_number['attachments'] = attachments
        print(f"Note Number: {note_number}")
        iw52NoteScreen.back()
        # break  # Remove this break to process all notes
    
    df = pd.DataFrame(notes)
    today_date = datetime.datetime.now().strftime("%Y%m%d")
    df.to_csv(f"notes_with_attachments_{today_date}_{ambiente_selecionado}.csv", index=False, encoding='utf-8-sig')

    iw52Screen.close()