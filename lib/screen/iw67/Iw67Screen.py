from win32com.client import CDispatch
from lib.screen.iw67.Iw67NotesMainScreen import (
    Iw67NotesMainScreen,
    Iw67NotesMainScreenInterface
)
from lib.types.Region import Region
import pandas as pd
class Iw67ScreenInterface:
    def openNotesScreen(self, region: Region) -> Iw67NotesMainScreenInterface: pass
    def back(self) -> None: pass
    def close(self) -> None: pass

class Iw67Screen(Iw67ScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def openNotesScreen(self, region: Region) -> Iw67NotesMainScreenInterface:
        if region == "SP":
            self._session.findById("wnd[0]/tbar[1]/btn[17]").press()
            self._session.findById("wnd[1]/usr/txtV-LOW").text = "CONSIRREGULAR"
            self._session.findById("wnd[1]/usr/txtENAME-LOW").text = ""
            self._session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
            self._session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
            self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
            self._session.findById("wnd[0]/tbar[1]/btn[8]").press()
        else:
            self._session.findById("wnd[0]/tbar[1]/btn[17]").press()
            self._session.findById("wnd[1]/usr/txtENAME-LOW").text = "204922"
            self._session.findById("wnd[1]/usr/txtENAME-LOW").setFocus()
            self._session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 6
            self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
            self._session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").setCurrentCell(5,"TEXT")
            self._session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"
            self._session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
            self._session.findById("wnd[0]/tbar[1]/btn[8]").press()
        
        return Iw67NotesMainScreen(self._session, self._connection)

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')

    def back(self) -> None:
        # self._session.findById("wnd[0]/tbar[0]/btn[3]").press()
        # self._session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        self._session.findById("wnd[0]").sendVKey(3)
