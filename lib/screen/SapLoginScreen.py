from lib.screen.SapHomeScreen import (
    SapHomeScreenInterface,
    SapHomeScreen
)
from win32com.client import CDispatch

from lib.types.Region import Region
from lib.types.SapServerName import SapServerName


class SapLoginScreenInterface:
    def login(self, user: str, password: str, region: Region, serverType: SapServerName) -> SapHomeScreenInterface: pass
    def close(self) -> None: pass

class SapMultipleLoginException(Exception):
    "Sap Multiple Login Exception: Rised when user try to login one more time"
    pass

class SapLoginScreen(SapLoginScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection
        # print(type(self._session))
    
    def setUser(self, user: str) -> None:
        self._session.findById("wnd[0]/usr/txtRSYST-BNAME").text = user

    def setPassword(self, password: str) -> None:
        self._session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password

    def setMandt(self, mandt: int) -> None:
        self._session.findById("wnd[0]/usr/txtRSYST-MANDT").text = mandt

    def pressLoginButton(self) -> None:
        self._session.findById("wnd[0]/tbar[0]/btn[1]").press()            
        self._session.findById("wnd[0]").sendVKey(0)
        try:
            self._session.findById("wnd[1]").sendVKey(0)
        except:
            pass

    def forceLogin(self) -> None:
        if self._session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1", False):
            self._session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").select()
            self._session.findById("wnd[1]/usr/radMULTI_LOGON_OPT1").setFocus()
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
            

        self._session.findById("wnd[0]").sendVKey(0)
        self._session.findById("wnd[0]").sendVKey(0)

    def login(self, user: str, password: str, region: Region, serverType: SapServerName) -> SapHomeScreenInterface:
        self._session.findById("wnd[0]").maximize()
        self.setUser(user)
        self.setPassword(password)
        if region == "ES" and serverType == "CP1":
            self.setMandt(600)
        self.pressLoginButton()
        self.forceLogin()
        try:
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except: pass
        try:
            self._session.findById("wnd[2]/tbar[0]/btn[0]").press()
        except: pass
        try:
            self._session.findById("wnd[2]/tbar[0]/btn[0]").press()
        except: pass
        try:
            self._session.findById("wnd[1]/tbar[0]/btn[12]").press()
        except: pass

        if "SAP Easy Access" in self._session.findById("wnd[0]").text:
            return SapHomeScreen(self._session, self._connection)
        else:
            raise Exception("Sap Login Exception. Change your credentials and try again.")

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')
