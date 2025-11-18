from lib.screen.SapLoginScreen import (
    SapLoginScreen,
    SapLoginScreenInterface
)
from lib.types.Region import Region
from lib.types.SapServerName import SapServerName
import win32com.client
from time import sleep
from os import system
import subprocess
import psutil


class SapLogonScreenInterface:
    def getSession(self, systemName: str) -> None: pass
    def isRunning(self) -> bool: pass
    def startNetweaver(self) -> None: pass
    def loadSystem(self, systemName: Region, systemType: SapServerName) -> SapLoginScreenInterface: pass
    def killNetweaver(self) -> None: pass


class SapLogonScreen(SapLogonScreenInterface):
    def getSession(
            self,
            connection: Region
    ) -> win32com.client.CDispatch:
        sleep(2)
        return connection.Children(0)

    def getConnection(
            self,
            systemName: Region,
            systemType: SapServerName
    ) -> win32com.client.CDispatch:

        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        if systemType in ["EP1", "EP2", "CP1"]:
            return application.OpenConnection(systemType)
        else:
            raise Exception("Tipo de servidor inválido")

    def loadSystem(self, systemName: Region, systemType: SapServerName) -> SapLoginScreenInterface:
        if not self.isRunning():
            self.startNetweaver()
        con = self.getConnection(systemName, systemType)
        session = self.getSession(con)
        return SapLoginScreen(session, con)

    def isRunning(self) -> bool:
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] == 'saplogon.exe':
                return True
        return False

    def startNetweaver(self) -> None:
        if not self.isRunning():
            subprocess.Popen(['start', 'saplogon.exe'], shell=True)
            sleep(3)

    def killNetweaver(self) -> None:
        if self.isRunning():
            system("taskkill /im saplogon.exe /f")
