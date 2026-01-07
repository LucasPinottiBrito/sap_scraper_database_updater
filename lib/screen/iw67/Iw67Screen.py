import datetime
from win32com.client import CDispatch
from lib.screen.iw67.Iw67NotesMainScreen import (
    Iw67NotesMainScreen,
    Iw67NotesMainScreenInterface
)
from lib.types.NoteTypes import NoteType
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

    def openNotesScreen(self, region: Region, noteType: NoteType = "SC/RC") -> Iw67NotesMainScreenInterface:
        match noteType:
            case "SC/RC":
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
                    self._session.findById("wnd[0]/usr/ctxtVARIANT").text = "/TRIAGEM_CI"
                    self._session.findById("wnd[0]/tbar[1]/btn[8]").press()
                
                return Iw67NotesMainScreen(self._session, self._connection)
            
            case "OVD":
                if region == 'SP':
                    self._session.findById("wnd[0]/usr/ctxtQMART-LOW").text = "ou"
                    self._session.findById("wnd[0]/usr/ctxtDATUV").text = "01012022"
                    self._session.findById("wnd[0]").sendVKey(0)
                    self._session.findById("wnd[0]/usr/ctxtM_PARVW").text = "ab"
                    self._session.findById("wnd[0]/usr/ctxtM_PARNR").text = "70234"
                    self._session.findById("wnd[0]/usr/ctxtVARIANT").text = "/OU_FELIPE"
                    self._session.findById("wnd[0]").sendVKey(8)
                else:
                    self._session.findById("wnd[0]/usr/btn%_QMART_%_APP_%-VALU_PUSH").press()
                    self._session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "OU"
                    self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    self._session.findById("wnd[0]/usr/btn%_MATXT_%_APP_%-VALU_PUSH").press()
                    self._session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "*DERI*"
                    self._session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").text = "*DDRC*"
                    self._session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,2]").text = "*DMRC*"
                    self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    self._session.findById("wnd[0]/usr/ctxtDATUV").text = "01012024"
                    self._session.findById("wnd[0]").sendVKey(0)
                    self._session.findById("wnd[0]/usr/ctxtDATUB").text = datetime.datetime.now().strftime("%d%m%Y")
                    self._session.findById("wnd[0]").sendVKey(0)
                    self._session.findById("wnd[0]/usr/ctxtVARIANT").text = "/DDRC OU PR"
                    self._session.findById("wnd[0]/tbar[1]/btn[8]").press()
                return Iw67NotesMainScreen(self._session, self._connection)
            
            case "PRC":
                self._session.findById("wnd[0]/usr/btn%_QMART_%_APP_%-VALU_PUSH").press()
                self._session.findById(
                    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "PR"
                self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
                self._session.findById("wnd[0]/usr/btn%_MATXT_%_APP_%-VALU_PUSH").press()
                self._session.findById(
                    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,0]").text = "*DERI*"
                self._session.findById(
                    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,1]").text = "*DDRC*"
                self._session.findById(
                    "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/txtRSCSEL_255-SLOW_I[1,2]").text = "*DMRC*"
                self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
                self._session.findById("wnd[0]/usr/ctxtDATUV").text = "01012024"
                self._session.findById("wnd[0]").sendVKey(0)
                self._session.findById("wnd[0]/usr/ctxtDATUB").text = datetime.datetime.now().strftime("%d%m%Y")
                self._session.findById("wnd[0]").sendVKey(0)
                # self._session.findById("wnd[0]/usr/ctxtVARIANT").text = "/DDRC OU PR"
                self._session.findById("wnd[0]/tbar[1]/btn[8]").press()
                try:
                    mygrid = self._session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
                    return Iw67NotesMainScreen(self._session, self._connection)
                except:
                    raise Exception("No notes found for PRC type.")
            
            case "REC":
                if region == 'SP':
                    # Abre os filtros
                    self._session.findById("wnd[0]/tbar[1]/btn[17]").press()
                    self._session.findById("wnd[1]/usr/txtENAME-LOW").text = "7119"
                    self._session.findById("wnd[1]/tbar[0]/btn[8]").press()
                    self._session.findById("wnd[0]").sendVKey(8)
                else:
                    self._session.findById("wnd[0]/usr/ctxtDATUV").text = ""
                    self._session.findById("wnd[0]/usr/ctxtMNCOD-LOW").text = "0009"
                    self._session.findById("wnd[0]/usr/ctxtPSTER-LOW").text = "01012024"
                    self._session.findById("wnd[0]/usr/ctxtPSTER-HIGH").text = datetime.datetime.now().strftime("%d%m%Y")
                    self._session.findById("wnd[0]").sendVKey(0)
                    self._session.findById("wnd[0]/usr/ctxtVARIANT").text = "/CI RECURSO"
                    self._session.findById("wnd[0]").sendVKey(8)
                return Iw67NotesMainScreen(self._session, self._connection)
            case _:
                raise NotImplementedError(f"Note type {noteType} not implemented yet.")

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')

    def back(self) -> None:
        try:
            self._session.findById("wnd[0]").sendVKey(3)
            return
        except:
            pass
        try:
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass
