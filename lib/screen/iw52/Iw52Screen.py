from lib.screen.iw52.Iw52NoteMainScreen import (
    Iw52NoteMainScreen,
    Iw52NoteMainScreenInterface
)
from win32com.client import CDispatch

class Iw52ScreenInterface:
    def openNote(self, number: str) -> Iw52NoteMainScreenInterface: pass
    def close(self) -> None: pass

class Iw52Screen(Iw52ScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def openNote(self, number: str) -> Iw52NoteMainScreenInterface:
        self._session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = number
        self._session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").caretPosition = 9
        self._session.findById("wnd[0]").sendVKey(0)
        self._session.findById("wnd[0]").sendVKey(0)
        if "1ª tela" in self._session.findById("wnd[0]").text:
            raise Exception(f"Note not found error. Not found {number} note in iw52")
        if self._session.findById("wnd[0]/tbar[1]/btn[17]", False):
            self._session.findById("wnd[0]/tbar[1]/btn[17]").press()
        return Iw52NoteMainScreen(self._session, self._connection)
    
    def close(self) -> None:
        self._connection.CloseSession('ses[0]')
