from win32com.client import CDispatch
from lib.types.Region import Region
import pandas as pd
class Iw67ScreenInterface:
    def openNotes(self, region: Region) -> None: pass
    def back(self) -> None: pass
    def close(self) -> None: pass

class Iw67Screen(Iw67ScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def openNotes(self, region: Region):
        try:
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
        except Exception as e:
            print(f"Error while opening notes: {e}")
            self.close()
            return
        
        try:
            self._session.findById("wnd[0]").maximize()
            # self._session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(-1,"")
            mygrid = self._session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            row_count = mygrid.RowCount
            table_content = []
            for i in range(row_count):
                # Nota, Criado em, Data da Nota, Concl.desejada, TextPrioridade, Texto de grupo de codificação, Texto de grupo de codes para medida, Texto de code medida, Cidade, Descrição
                nota = mygrid.GetCellValue(i, "QMNUM")
                criado_em = mygrid.GetCellValue(i, "ERDAT")
                conclu_desejada = mygrid.GetCellValue(i, "LTRMN")
                textPrioridade = mygrid.GetCellValue(i, "PRIOKX")
                texto_de_grupo_de_codificacao = mygrid.GetCellValue(i, "KTXTCD")
                texto_de_code_medida = mygrid.GetCellValue(i, "SMCODETEXT")
                cidade = mygrid.GetCellValue(i, "CITY1")
                descricao = mygrid.GetCellValue(i, "QMTXT")
                table_content.append({
                    "nota": nota,
                    "criado_em": criado_em,
                    "conclu_desejada": conclu_desejada,
                    "textPrioridade": textPrioridade,
                    "texto_de_grupo_de_codificacao": texto_de_grupo_de_codificacao,
                    "texto_de_code_medida": texto_de_code_medida,
                    "cidade": cidade,
                    "descricao": descricao
                })
                print(f"Nota: {nota}, Criado em: {criado_em}, Concl.desejada: {conclu_desejada}, TextPrioridade: {textPrioridade}, Texto de grupo de codificação: {texto_de_grupo_de_codificacao}, Texto de code medida: {texto_de_code_medida}, Cidade: {cidade}, Descrição: {descricao}")

            pd.DataFrame(table_content).to_csv("notas_iw67.csv", index=False, sep=";")
        
        except Exception as e:
            print(f"Error while extracting notes data: {e}")    
            
        self.close()

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')

    def back(self) -> None:
        self._session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self._session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
