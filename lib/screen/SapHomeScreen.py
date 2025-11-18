from lib.screen.cic0.cic0Screen import (cic0Screen, cic0ScreenInterface)
from lib.screen.iw52.Iw52Screen import (Iw52Screen, Iw52ScreenInterface)
from lib.screen.iw67.Iw67Screen import (Iw67Screen, Iw67ScreenInterface)
from win32com.client import CDispatch
from typing import Literal

class SapHomeScreenInterface:
    def openTransaction(self, transaction: Literal["iw52", "cic0", "iw67"]):
        interfaceDict = {
            "iw52": Iw52ScreenInterface,
            "iw67": Iw67ScreenInterface,
            "cic0": cic0ScreenInterface
        }
        return interfaceDict.get(transaction)
    
    def close(self) -> None: pass

class SapHomeScreen:
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def openTransaction(self, transaction: Literal["iw52", "iw67", "cic0"]):
        self._session.findById("wnd[0]").sendVKey(0)
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/tbar[0]/okcd").text = transaction
        self._session.findById("wnd[0]").sendVKey(0)
        transactionDict = {
            "iw52": Iw52Screen,
            "cic0": cic0Screen,
            "iw67": Iw67Screen
        }
        return transactionDict.get(transaction)(self._session, self._connection)

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')
