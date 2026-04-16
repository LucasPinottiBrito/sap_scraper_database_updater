from datetime import datetime
from typing import List, Literal
from win32com.client import CDispatch
from lib.types.NoteDetails import NoteDetails
import requests
import os

dicionario_medidas_para_transferencia_sem_ci = {
    "AL13": "AL21",
    "AL14": "AL22",
    "AL15": "AL23",
    "AL16": "AL24",
    "AL17": "AL21",
    "AL18": "AL22",
    "AL19": "AL23",
    "AL20": "AL24"
}

dicionario_medidas_para_transferencia_mmgd = {
    "AL13": "AL42",
    "AL14": "AL37",
    "AL15": "AL42",
    "AL16": "AL37",
    "AL17": "AL42",
    "AL18": "AL37",
    "AL19": "AL42",
    "AL20": "AL37"
}

dicionario_medidas_procedencia = {
    "3": "RPCA",
    "2": "RPEM",
    "6": "RPFX",
    "7": "RPIN",
    "4": "RPSC",
    "1": "RPSM",
    "5": "RPTE"
}

texto_resultado_procedencia_default = """Olá! Informamos que sua solicitação de alteração de responsabilidade foi
atendida com sucesso.
Agradecemos o seu contato e se precisar de esclarecimentos ou novas
informações, entre em contato através de nosso 0800-721-4322 de segunda
a sexta-feira das 8:30 às 12h e das 13:30 às 16:30, exceto feriados
nacionais, ou procure uma de nossas Agências mais próxima, veja os
locais em
https://www.edp.com.br/canais-de-atendimento/atendimento-presencial,
para outros serviços acesse nossa agência virtual edponline.com.br.
"""

texto_medida_procedencia_default = "Solicitação atendida."
texto_resultado_improcedencia_default = "Solicitação improcedente."

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


class InfoContatoInterface:
    contato_email: str
    contato_sms: str
    cod_contato: str

    def __init__(self, contato_email: str, contato_sms: str, cod_contato: str) -> None:
        self.contato_email = contato_email
        self.contato_sms = contato_sms
        self.cod_contato = cod_contato


class Iw52NoteMainScreenInterface:
    def isOpen(self) -> bool: pass
    def get_attachments(self, download_files: bool, folder_path_to_download: str) -> list[str]: pass
    def getNoteDetails(self) -> NoteDetails: pass
    def proceder(self, texto_resultado: str, texto_medida: str) -> None: pass
    def improceder(self, texto_resultado: str, texto_medida: str) -> None: pass
    def transferir_para_sem_ci(self) -> None: pass
    def transferir_para_mmgd(self) -> None: pass
    def encerrar_nota(self) -> None: pass
    def procederAlteracaoResponsabilidade(self, texto_resultado: str, texto_medida: str) -> None: pass
    def improcederAlteracaoResponsabilidade(self, texto_resultado: str, texto_medida: str) -> None: pass
    def transferirNotaSemCI(self) -> None: pass
    def transferirNotaMMGD(self) -> None: pass
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

    def _salvar_alteracoes(self) -> None:
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/tbar[0]/btn[11]").press()
        self._session.findById("wnd[0]").sendVKey(0)

    def _formatar_texto_resultado(self, texto: str) -> list[str]:
        texto = texto.replace("\n", " ").replace("\r", " ")
        return [texto[i:i+72] for i in range(0, len(texto), 72)]

    def _adicionar_causa(self, cod_causa: str, desc_causa, cod_subcausa: str, desc_subcausa: str) -> None:
        self._session.findById("wnd[0]").maximize()
        self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21").select()
        self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtZCCSTWM_R12_DC-COD_CAUSA").text = cod_causa
        self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtZCCSTWM_R12_DC-DESC_CAUSA").text = desc_causa
        self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtZCCSTWM_R12_DC-COD_SUBCAUSA").text = cod_subcausa
        self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtZCCSTWM_R12_DC-DESC_SUBCAUSA").text = desc_subcausa
        self._session.findById("wnd[0]").sendVKey(0)

    def _adicionar_resultado(self, cod_resultado: str, cod_procedencia: str, contato: str, texto_resultado: str) -> None:
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12").select()
        i = 0
        while True:
            self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNGRP[1,{i}]").setFocus()
            if self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNGRP[1,{i}]").text == "":
                break
            i += 1
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNGRP[1,{i}]").text = cod_resultado
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/ctxtVIQMMA-MNCOD[2,{i}]").text = cod_procedencia
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/txtVIQMMA-MATXT[4,{i}]").text = contato
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/txtVIQMMA-MATXT[4,{i}]").setFocus()
        self._session.findById("wnd[0]").sendVKey(0)

        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB12/ssubSUB_GROUP_10:SAPLIQS0:7130/tblSAPLIQS0AKTIONEN_VIEWER/btnQMICON-LTAKTION[5,0]").press()
        formatted_result = self._formatar_texto_resultado(texto_resultado)
        for i, texto in enumerate(formatted_result):
            self._session.findById(f"wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,{i+2}]").text = texto
        self._session.findById("wnd[0]").sendVKey(0)
        self._session.findById("wnd[0]/tbar[0]/btn[3]").press()

    def _adicionar_medida(self, grp_code: str, cod_medida: str, texto_medida: str = "", texto_resultado: str = "") -> None:
        i = 0
        while True:
            if self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,{i}]").text == "":
                break
            i += 1
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select()
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[1,{i}]").text = grp_code
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNCOD[2,{i}]").text = cod_medida or ""
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-MATXT[4,{i}]").text = texto_medida
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-MATXT[4,{i}]").setFocus()
        self._session.findById("wnd[0]").sendVKey(0)
        if texto_resultado.strip() != "":
            self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/btnQMICON-LTMASS[5,{i}]").press()
            formatted_result = self._formatar_texto_resultado(texto_resultado)
            for i, texto in enumerate(formatted_result):
                self._session.findById(f"wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,{i+2}]").text = texto
            self._session.findById("wnd[0]").sendVKey(0)
            self._session.findById("wnd[0]/tbar[0]/btn[3]").press()

    def _adicionar_fim_avaria(
        self,
        data_fim: str = datetime.now().strftime("%d.%m.%Y"),
        hora_fim: str = datetime.now().strftime("%H:%M:%S")
    ):
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01").select()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7520/ctxtVIQMEL-AUSBS").text = data_fim
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7520/ctxtVIQMEL-AUZTB").text = hora_fim
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7520/ctxtVIQMEL-AUZTB").setFocus()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7520/ctxtVIQMEL-AUZTB").caretPosition = 8
        self._session.findById("wnd[0]").sendVKey(0)

    def _encerrar_ultima_medida_existente(self, texto_medida: str = "", descricao_medida: str = ""):
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select()
        i = 0
        while True:
            if self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,{i}]").text == "":
                break
            i += 1
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-MATXT[4,{i-1}]").SetFocus()
        self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-MATXT[4,{i-1}]").text = descricao_medida
        if texto_medida.strip() != "":
            self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/btnQMICON-LTMASS[5,{i-1}]").press()
            formatted_result = self._formatar_texto_resultado(texto_medida)
            for index, texto in enumerate(formatted_result):
                self._session.findById(f"wnd[0]/usr/tblSAPLSTXXEDITAREA/txtRSTXT-TXLINE[2,{index+2}]").text = texto
            self._session.findById("wnd[0]").sendVKey(0)
            self._session.findById("wnd[0]/tbar[0]/btn[3]").press()

        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER").getAbsoluteRow(i-1).selected = True
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/btnFC_ERLEDIGT").press()

    def _get_ultima_medida_existente(self) -> str:
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11").select()
        i = 0
        while True:
            if self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/txtVIQMSM-QSMNUM[0,{i}]").text == "":
                break
            i += 1
        return self._session.findById(f"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB11/ssubSUB_GROUP_10:SAPLIQS0:7120/tblSAPLIQS0MASSNAHMEN_VIEWER/ctxtVIQMSM-MNGRP[4,{i-1}]").text

    def _get_info_contato(self) -> InfoContatoInterface:
        self._session.findById("wnd[0]").maximize()
        self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21").select()
        contato_email = self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/txtW_WM_R12_DC-EMAIL").text
        conato_sms = self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/txtW_WM_R12_DC-TEL").text
        cod_contato = self._session.findById(r"wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB21/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7900/subUSER0001:SAPLXQQM:0590/ctxtW_WM_R12_DC-COD_RESP").text
        return InfoContatoInterface(
            contato_email=contato_email,
            contato_sms=conato_sms,
            cod_contato=cod_contato
        )

    def _set_status_nota(self, status: Literal["ABER", "ANA", "VERI", "ENCP", "ENCI"]):
        status_index = ["ABER", "ANA", "VERI", "ENCP", "ENCI"].index(status)
        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01").select()
        self._session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/btnANWENDERSTATUS").press()
        self._session.findById(f"wnd[1]/usr/sub:SAPLBSVA:0201[0]/radJ_STMAINT-ANWS[{status_index},0]").select()
        self._session.findById("wnd[1]/tbar[0]/btn[0]").press()

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
            try:
                filename = f'{str(i)}_{self._session.findById("wnd[1]/usr/txtSOS17-S_URL_DESC").text}'
                url = self._session.findById("wnd[1]/usr/txtSOS17-S_URL_KEY").text
            except:
                self._session.findById("wnd[2]").close()
                filename = f'error_attachment'
                url = ''
                print(f"Error retrieving attachment details for row {i}. Skipping download.")
            
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
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01").select()
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7515/subANSPRECH:SAPLIPAR:0700/tabsTSTRIP_700/tabpKUND/ssubTSTRIP_SCREEN:SAPLIPAR:0130/subADRESSE:SAPLIPAR:0150/txtDIADR-NAME1").setFocus()
                nome_cliente = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7515/subANSPRECH:SAPLIPAR:0700/tabsTSTRIP_700/tabpKUND/ssubTSTRIP_SCREEN:SAPLIPAR:0130/subADRESSE:SAPLIPAR:0150/txtDIADR-NAME1").text
                
                return {
                    "number": self._noteNumber,
                    "description_text": descrption_text,
                    "contato_email": contato_email,
                    "contato_sms": conato_sms,
                    "cod_contato": cod_contato,
                    "inst": inst,
                    "nome_cliente": nome_cliente
                }
            case "CI":
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02").select()
                self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7515/subANSPRECH:SAPLIPAR:0700/tabsTSTRIP_700/tabpKUND/ssubTSTRIP_SCREEN:SAPLIPAR:0130/subADRESSE:SAPLIPAR:0150/txtDIADR-NAME1").setFocus()
                nome_cliente = self._session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\TAB02/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_1:SAPLIQS0:7515/subANSPRECH:SAPLIPAR:0700/tabsTSTRIP_700/tabpKUND/ssubTSTRIP_SCREEN:SAPLIPAR:0130/subADRESSE:SAPLIPAR:0150/txtDIADR-NAME1").text
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
                    "inst": inst,
                    "nome_cliente": nome_cliente
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
            self._session.findById("wnd[0]").sendVKey(12)
        except:
            pass
        try:
            self._session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
        except:
            pass
        try:
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except:
            pass
        return

    def procederAlteracaoResponsabilidade(
        self,
        texto_resultado: str = texto_resultado_procedencia_default,
        texto_medida: str = texto_medida_procedencia_default
    ) -> None:
        self._adicionar_causa(
            cod_causa="CIR006",
            desc_causa="TROCA DE TITULARIDADE",
            cod_subcausa="CR0604",
            desc_subcausa="INSENÇÃO DE DEBITOS CONSUMO IRREGULAR",
        )
        info_contato = self._get_info_contato()
        cod_resposta = dicionario_medidas_procedencia.get(info_contato.cod_contato.strip(), "RPSC")
        contato = info_contato.contato_email if info_contato.cod_contato.strip() == "2" else info_contato.contato_sms
        self._adicionar_resultado("PROC", cod_resposta, contato, texto_resultado)
        self._encerrar_ultima_medida_existente(texto_medida)
        self._adicionar_fim_avaria()
        self._set_status_nota("ENCP")
        self.terminate()

    def proceder(self, texto_resultado: str, texto_medida: str) -> None:
        self.procederAlteracaoResponsabilidade(
            texto_resultado=texto_resultado,
            texto_medida=texto_medida,
        )

    def improcederAlteracaoResponsabilidade(
        self,
        texto_resultado: str = texto_resultado_improcedencia_default,
        texto_medida: str = ""
    ) -> None:
        self._adicionar_causa(
            cod_causa="CIR006",
            desc_causa="TROCA DE TITULARIDADE",
            cod_subcausa="CR0604",
            desc_subcausa="INSENÇÃO DE DEBITOS CONSUMO IRREGULAR",
        )
        info_contato = self._get_info_contato()
        cod_resposta = dicionario_medidas_procedencia.get(info_contato.cod_contato.strip(), "RPSC")
        contato = info_contato.contato_email if info_contato.cod_contato.strip() == "2" else info_contato.contato_sms
        self._adicionar_resultado("IMPR", cod_resposta, contato, texto_resultado)
        self._encerrar_ultima_medida_existente(texto_medida)
        self._adicionar_fim_avaria()
        self._set_status_nota("ENCI")
        self.terminate()

    def improceder(self, texto_resultado: str, texto_medida: str) -> None:
        self.improcederAlteracaoResponsabilidade(
            texto_resultado=texto_resultado,
            texto_medida=texto_medida,
        )

    def transferirNotaSemCI(self) -> None:
        try:
            self._encerrar_ultima_medida_existente("SEM CI")
            cod_medida = dicionario_medidas_para_transferencia_sem_ci.get(self._get_ultima_medida_existente())
            self._adicionar_medida("NAALT", cod_medida=cod_medida)
            self._salvar_alteracoes()
        except Exception as e:
            raise Exception(f"Erro ao transferir nota (SEM CI): {e}") from e

    def transferir_para_sem_ci(self) -> None:
        self.transferirNotaSemCI()

    def transferirNotaMMGD(self) -> None:
        try:
            self._encerrar_ultima_medida_existente("MMGD")
            cod_medida = dicionario_medidas_para_transferencia_mmgd.get(self._get_ultima_medida_existente())
            self._adicionar_medida("NAALT", cod_medida=cod_medida)
            self._salvar_alteracoes()
        except Exception as e:
            raise Exception(f"Erro ao transferir nota MMGD: {e}") from e

    def transferir_para_mmgd(self) -> None:
        self.transferirNotaMMGD()

    def encerrar_nota(self) -> None:
        self.terminate()

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')
