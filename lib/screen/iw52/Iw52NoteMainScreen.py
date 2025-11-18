from win32com.client import CDispatch
from lib.types.Note import Note

class Iw52NoteMainScreenInterface:
    def getNote(self) -> Note: pass
    def isOpen(self) -> bool: pass
    def reOpenNote(self) -> None: pass
    def terminate(self) -> None: pass
    def close(self) -> None: pass

class Iw52NoteMainScreen(Iw52NoteMainScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def __getComponentById(self, id: str) -> str | None:
        componentInput = self._session.findById(id)
        if componentInput:
            return componentInput.text

    def getNoteNumber(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-QMNUM")

    def getNoteType(self) -> str | None:
         return self.__getComponentById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/ctxtVIQMEL-QMART")

    def getNoteText(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-QMTXT")

    def getSystemStatus(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtRIWO00-STTXT")

    def getUserStatus(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtRIWO00-ASTXT")

    def getEquipamentNumber(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-EQUNR")

    def getSerieNumber(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-SERIALNR")

    def getMaterial(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7322/subOBJEKT:SAPLIWO1:1200/ctxtRIWO1-MATNR")

    def getAffectedAnlage(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_2:SAPLIQS0:7318/ctxtVIQMEL-BTPLN")

    def getDamageSympton(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/ctxtVIQMFE-FEGRP")

    def getDamageCode(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/ctxtVIQMFE-FECOD")

    def getItemText(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_3:SAPLIQS0:7324/txtVIQMFE-FETXT")

    def isOpen(self) -> bool:
        try:
            button = self._session.findByiD("wnd[0]/tbar[1]/btn[17]", False)
        except Exception as err:
            print(err)
        if button:
            return False
        return True

    def get_attachments(self) -> list[str]:
        self._session.findById("wnd[0]").maximize
        self._session.findById("wnd[0]/titl/shellcont/shell").pressButton("%GOS_TOOLBOX")
        self._session.findById("wnd[1]/usr/tblSAPLSWUGOBJECT_CONTROL/txtSWLOBJTDYN-DESCRIPT[0,0]").caretPosition = 10
        self._session.findById("wnd[1]").sendVKey(2)
        self._session.findById("wnd[0]/shellcont/shell").pressButton("VIEW_ATTA")
        grid = self._session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell")
        row_count = grid.RowCount
        attachments = []
        for i in range(row_count):
            self._session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").selectedRows = str(i)
            self._session.findById("wnd[1]/usr/cntlCONTAINER_0100/shellcont/shell").doubleClickCurrentCell()


    def getNote(self) -> Note:
        self._session.findById("wnd[0]").maximize
        return {
            "number": self.getNoteNumber(),
            "type": self.getNoteType(),
            "text": self.getNoteText(),
            "systemStatus": self.getSystemStatus(),
            "userStatus": self.getUserStatus(),
            "equipamentNumber": self.getEquipamentNumber(),
            "serieNumber": self.getSerieNumber(),
            "material": self.getMaterial(),
            "affectedAnlage": self.getAffectedAnlage(),
            "damageSympton": self.getDamageSympton(),
            "damageCode": self.getDamageCode(),
            "itemText": self.getItemText(),
            "isOpen": self.isOpen()
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
    
    def close(self) -> None:
        self._connection.CloseSession('ses[0]')