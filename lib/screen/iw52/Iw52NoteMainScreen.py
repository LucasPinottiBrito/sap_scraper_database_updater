from typing import List, Literal
from win32com.client import CDispatch
from lib.types.NoteDetails import NoteDetails
import requests
import os

def download_attachment(url: str, folder_path, file_path: str) -> None:
    response = requests.get(url, stream=True)
    if response.status_code == 200:
        with open(file_path, 'wb') as file:
            file.write(response.content)
        
        if os.path.exists(folder_path + os.path.basename(file_path)):
            os.remove(folder_path + os.path.basename(file_path))
        
        os.rename(file_path, folder_path + os.path.basename(file_path))
    else:
        print(f"Failed to download attachment from {url}. Status code: {response.status_code}")

class Iw52NoteMainScreenInterface:
    def isOpen(self) -> bool: pass
    def get_attachments(self, download_files: bool, folder_path_to_download: str) -> list[str]: pass
    def getNoteDetails(self) -> NoteDetails: pass
    def reOpenNote(self) -> None: pass
    def terminate(self) -> None: pass
    def back(self) -> None: pass
    def close(self) -> None: pass

class Iw52NoteMainScreen(Iw52NoteMainScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    _noteNumber: str

    def __init__(self, session: CDispatch, connection: CDispatch, noteNumber: str) -> None:
        self._session = session
        self._connection = connection
        self._noteNumber = noteNumber

    def isOpen(self) -> bool:
        try:
            button = self._session.findByiD("wnd[0]/tbar[1]/btn[17]", False)
        except Exception as err:
            print(err)
        if button:
            return False
        return True

    def get_attachments(self, download_files: bool, folder_path_to_download = '') -> List[str]:
        self._session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
        self._session.findById("wnd[1]/usr/tblSAPLSWUGOBJECT_CONTROL").getAbsoluteRow(0).selected = True
        self._session.findById("wnd[1]").sendVKey(0)
        self._session.findById("wnd[0]/shellcont/shell").pressButton("VIEW_ATTA")
        grid = self._session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
        row_count = grid.RowCount
        url_list = []
        if row_count == 0:
            self._session.findById("wnd[1]").close()
            self._session.findById("wnd[0]/shellcont").close()
            return []
        for i in range(row_count):
            grid.selectedRows = str(i)
            grid.pressToolbarButton("%ATTA_EDIT")
            filename = f'{str(i)}_{self._session.findById("wnd[1]/usr/txtSOS17-S_URL_DESC").text}'
            url = self._session.findById("wnd[1]/usr/txtSOS17-S_URL_KEY").text
            
            folder_path = folder_path_to_download or f"./attachments/{self._noteNumber}/"
            if download_files:
                if not os.path.exists(folder_path): #create folder if not exists
                    os.makedirs(folder_path)
                download_attachment(url, folder_path, filename)
            
            self._session.findById("wnd[1]/tbar[0]/btn[12]").press()
            url_list.append(url)
        
        self._session.findById("wnd[1]/tbar[0]/btn[12]").press()
        self._session.findById("wnd[2]/usr/btnSPOP-OPTION1").press()
        self._session.findById("wnd[0]/shellcont").close()
        self._session.findById("wnd[0]").maximize()
        return url_list


    def getNoteDetails(self, note_model: Literal["CI", "NA"] = "NA") -> NoteDetails:
        match note_model:
            case "NA":
                self._session.findById("wnd[0]").maximize()
                descrption_text = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7715/cntlTEXT/shellcont/shell").text
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21").select()
                contato_email = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/txtW_WM_R12_DC-EMAIL").text
                conato_sms = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/txtW_WM_R12_DC-TEL").text
                cod_contato = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtW_WM_R12_DC-COD_RESP").text
                inst = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtW_WM_R12_DC-ANLAGE").text
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select()

                return {
                    "number": self._noteNumber,
                    "description_text": descrption_text,
                    "contato_email": contato_email,
                    "contato_sms": conato_sms,
                    "cod_contato": cod_contato,
                    "inst": inst
                }
            case "CI":
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09").select()
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB09/ssubSUB_GROUP_10:SAPLIQS0:7217/subSUBSCREEN_1:SAPLIQS0:7790/subUSER0001:SAPLXQQM:0101/btnBTTOI").press()
                inst = self._session.findById("wnd[0]/usr/ctxtZCCSTBI_FR_TOI-ANLAGE").text
                self._session.findById("wnd[0]/tbar[0]/btn[3]").press()
                try:
                    self._session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press()
                except:
                    pass
                return {
                    "number": self._noteNumber,
                    "description_text": "",
                    "contato_email": "",
                    "contato_sms": "",
                    "cod_contato": "",
                    "inst": inst
                }
    
    def reOpenNote(self) -> None:
        self._session.findById("wnd[0]").maximize
        self._session.findById("wnd[0]/tbar[1]/btn[17]").press()

    def terminate(self) -> None:
        self._session.findById("wnd[0]").maximize
        self._session.findById("wnd[0]/tbar[1]/btn[16]").press()
        self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            self._session.findById("wnd[2]/tbar[0]/btn[0]").press()
        except:
            pass
        
    def back(self) -> None:
        try:
            self._session.findById("wnd[0]").sendVKey(3)
        except:
            pass
        try:
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass
        try:
            self._session.findById("wnd[1]/usr/btnSPOP-OPTION2").press()
        except:
            pass
        return

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')