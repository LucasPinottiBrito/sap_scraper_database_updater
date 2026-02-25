from win32com.client import CDispatch

note_type_dictionary = {
    "REC": "Recurso",
    "PRC": "Procon",
    "SC/RC": "Solicitacao",
    "OVD": "Ouvidoria"
}

class Iw67NotesMainScreenInterface:
    def getNotes(self) -> list[dict]: pass
    def isOpen(self) -> bool: pass
    def close(self) -> None: pass
    def back(self) -> None: pass

class Iw67NotesMainScreen(Iw67NotesMainScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def __getComponentById(self, id: str) -> str | None:
        componentInput = self._session.findById(id)
        if componentInput:
            return componentInput.text

    def getMainGrid(self) -> str | None:
        return self.__getComponentById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/txtVIQMEL-QMNUM")

    def isOpen(self) -> bool:
        try:
            button = self._session.findByiD("wnd[0]/tbar[1]/btn[17]", False)
        except Exception as err:
            print(err)
        if button:
            return False
        return True

    def getNotes(self, note_type: str = None) -> list[dict]:
        try:
            self._session.findById("wnd[0]").maximize()
            mygrid = self._session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")
            row_count = mygrid.RowCount
            
            table_content = []
            
            if note_type:
                tipo_nota = note_type_dictionary.get(note_type, "")
            else:
                tipo_nota = ""

            for i in range(row_count):
                nota = mygrid.GetCellValue(i, "QMNUM")
                criado_em = mygrid.GetCellValue(i, "ERDAT")

                try: 
                    data_conclusao = mygrid.GetCellValue(i, "PETER")
                    data_conclusao = data_conclusao.replace(".", "-").replace("/", "-")
                except: 
                    data_conclusao = ""
                
                try: 
                    conclu_desejada = mygrid.GetCellValue(i, "LTRMN")
                    conclu_desejada = conclu_desejada.replace(".", "-").replace("/", "-")
                except: 
                    conclu_desejada = ""
        
                try: textPrioridade = mygrid.GetCellValue(i, "PRIOKX")
                except: textPrioridade = ""

                try: texto_de_grupo_de_codificacao = mygrid.GetCellValue(i, "KTXTCD")
                except: texto_de_grupo_de_codificacao = ""
                
                try: texto_de_code_medida = mygrid.GetCellValue(i, "SMCODETEXT")
                except: texto_de_code_medida = ""

                try: cidade = mygrid.GetCellValue(i, "CITY1")
                except: cidade = ""

                try: descricao = mygrid.GetCellValue(i, "QMTXT")
                except: descricao = ""

                try: pn = mygrid.GetCellValue(i, "KUNUM")
                except: pn = ""

                table_content.append({
                    "note_number": nota,
                    "created_at": criado_em.replace(".", "-").replace("/", "-") if criado_em != "" else "",
                    "conclusion_date": (data_conclusao or conclu_desejada),
                    "priority_text": textPrioridade,
                    "group": texto_de_grupo_de_codificacao,
                    "code_text": texto_de_code_medida,
                    "city": cidade,
                    "description": descricao,
                    "business_partner_id": pn,
                    "note_type": tipo_nota
                })
                print(f"Nota: {nota}, Criado em: {criado_em}, Concl.desejada: {conclu_desejada}, TextPrioridade: {textPrioridade}, Texto de grupo de codificação: {texto_de_grupo_de_codificacao}, Texto de code medida: {texto_de_code_medida}, Cidade: {cidade}, Descrição: {descricao}")
            self.back()
            return table_content
        
        except Exception as e:
            print(f"Error while extracting notes data: {e}")
            self.back()    
            return []
    
    def close(self) -> None:
        self._connection.CloseSession('ses[0]')

    def back(self) -> None:
        self._session.findById("wnd[0]").sendVKey(3)