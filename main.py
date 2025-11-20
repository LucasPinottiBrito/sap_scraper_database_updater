from lib.screen.SapLogonScreen import SapLogonScreen
import pandas as pd
import time

ambiente = ""
usuario = ""
senha = ""

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
    for note_number in notes:
        iw52Screen = home.openTransaction("iw52")
        print(f"Processing Note Number: {note_number.get('nota')}")
        iw52NoteScreen = iw52Screen.openNote(note_number.get("nota"))
        attachments = iw52NoteScreen.get_attachments()
        print(f"Note Number: {note_number}, Attachments: {attachments}")
        break  # Remove this break to process all notes