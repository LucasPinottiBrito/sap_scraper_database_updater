from win32com.client import CDispatch

class cic0ScreenInterface:
    def openInstalationInfo(self, inst: str) -> str: pass
    def switchClass(self, inst: str) -> str: pass
    def back(self) -> None: pass
    def close(self) -> None: pass

class cic0Screen(cic0ScreenInterface):
    _session: CDispatch
    _connection: CDispatch
    def __init__(self, session: CDispatch, connection: CDispatch) -> None:
        self._session = session
        self._connection = connection

    def createVL(self, inst: str) -> str:
        try:
            self._session.findById("wnd[1]/usr/lbl[43,3]").SetFocus()
            self._session.findById("wnd[1]/usr/lbl[43,3]").caretPosition = 9
            self._session.findById("wnd[1]").sendVKey(2)
            self._session.findById("wnd[0]/usr/subAREA05:SAPLCRM_CIC_NAV_AREA:0101/tabsTABSTRIP/tabpNA_TAB1/ssubSUB_TAB:SAPLCRM_CIC_NAV_AREA:0105/subNAV_CUST_SUB:SAPLBUS_LOCATOR:3102/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3211/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell[1]/shellcont[1]/shell[1]").hierarchyHeaderWidth = 258
        except:
            pass

        try:
           self._session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/3[0,3]/lbl[1,3]").SetFocus()
           self._session.findById("wnd[1]/usr/sub/1[0,0]/sub/1/2[0,0]/sub/1/2/3[0,3]/lbl[1,3]").caretPosition = 13
           self._session.findById("wnd[1]").sendVKey(2)
        except:
            pass

        self._session.findById("wnd[0]").maximize()
        self._session.findById("wnd[0]/usr/subAREA06:SAPLCCM21:0101/tabsTABSTRIP/tabpAA_TAB1/ssubSUB_TAB:SAPLCCM21:0105/ssubCCM21_CUST_SUB:SAPLCCM1:0501/subSEARCH_DISPLAY:SAPLCRM_CIC_BP_SUB:6006/cntlHTML_CONTROL_BPSEARCH/shellcont/shell").sapEvent("","BP1_PARTNER=&BP1_MCNAME1=&BP1_MCNAME2=&C_TAXNUM=&BP2_PARTNER=&BP2_MCNAME1=&BP2_MCNAME2=&TIME_ZONE=&A_VKONT=&SEARCHSTRING=&ADRTYP=BP1&STREET=&HOUSE_NUM1=&CITY1=&POST_CODE1=&COUNTRY=&TELEPHONE=&TEL_EXT=&FAX=&FAX_EXT=&EMAIL=&LN_CHANGE=<SPAN><IMG src%3DS_F_SAVE.GIF align%3Dcenter> </SPAN>&FIND_BUTTON=<SPAN><IMG src%3DS_B_SRCH.gif align%3Dcenter></SPAN>&CR_TRGT=<SPAN><IMG src%3DS_B_CREA.gif align%3Dcenter></SPAN>","sapevent:ISU_FIND")
        self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2").select()
        self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-I_ANLAGE").text = str(inst)
        self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-I_ANLAGE").setFocus()
        self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-I_ANLAGE").caretPosition = 7
        self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
        try:
            self._session.findById("wnd[0]/tbar[1]/btn[5]").press()
            self._session.findById("wnd[0]").sendVKey(12)
        except:
            pass

        try:
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self._session.findById("wnd[1]").close()
            self._session.findById("wnd[0]/tbar[1]/btn[6]").press()
            notaGerada = "Instalação nunca teve contrato : Impossível gerar VL"
            return notaGerada
        except:
            self._session.findById("wnd[0]/usr/subAREA04:SAPLCRM_CIC_SLIM_ACTION_BOX:0100/cntlCRMCICSLIMABCONTAINER/shellcont/shell").pressContextButton("0008")
            self._session.findById("wnd[0]/usr/subAREA04:SAPLCRM_CIC_SLIM_ACTION_BOX:0100/cntlCRMCICSLIMABCONTAINER/shellcont/shell").selectContextMenuItem("NSSV")
            self._session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
            self._session.findById("wnd[1]/usr/btnBUTTON_1").press()
            self._session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = inst
            self._session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").caretPosition = 7
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self._session.findById("wnd[1]/usr/btnBUTTON_1").press()

            try:
                self._session.findById("wnd[1]/tbar[0]/btn[5]").press()
                self._session.findById("wnd[0]/tbar[1]/btn[6]").press()
                notaGerada = "Instalação é de Média Tensão : Impossível gerar VL"
                return notaGerada

            except:
                self._session.findById("wnd[1]/usr/btnDY_VAROPTION1").press()

            self._session.findById("wnd[1]/usr/cntlCUSTOM_CONTAINER/shell").text = "Nota de revisita em cliente autuado. Favor registrar foto padrão, fechada do imóvel e foto do comprovante de entrega."  # COLAR TEXTO DA NOTA AQUI
            self._session.findById("wnd[1]/tbar[0]/btn[16]").press()
            notaGerada = self._session.findById("wnd[1]/usr/txtD0100-TEXT_FIELD_02").text
            self._session.findById("wnd[1]/tbar[0]/btn[5]").press()
            self._session.findById("wnd[1]/usr/cntlTEXTEDITOR1/shellcont/shell").text = "Nota de revisita em cliente autuado. Favor registrar foto padrão, fechada do imóvel e foto do comprovante de entrega."  # COLAR TEXTO DA NOTA AQUI
            self._session.findById("wnd[1]/usr/cntlTEXTEDITOR1/shellcont/shell").setSelectionIndexes(1, 1)
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self._session.findById("wnd[1]/tbar[0]/btn[5]").press()
            self._session.findById("wnd[0]/tbar[1]/btn[6]").press()
            return notaGerada
        
    def switchClass(self, inst: str) -> str:
        try:
            self._session.findById("wnd[1]/usr/lbl[22,6]").setFocus()
            self._session.findById("wnd[1]/usr/lbl[22,6]").caretPosition = 8
            self._session.findById("wnd[1]").sendVKey(2)
            self._session.findById("wnd[0]/usr/subAREA05:SAPLCRM_CIC_NAV_AREA:0101/tabsTABSTRIP/tabpNA_TAB1/ssubSUB_TAB:SAPLCRM_CIC_NAV_AREA:0105/subNAV_CUST_SUB:SAPLBUS_LOCATOR:3102/ssubSCREEN_3100_TABSTRIP_AREA:SAPLBUS_LOCATOR:3200/subSCREEN_3200_SEARCH_AREA:SAPLBUS_LOCATOR:3214/subSCREEN_3200_RESULT_AREA:SAPLBUS_LOCATOR:3250/cntlSCREEN_3210_CONTAINER/shellcont/shell[1]/shellcont[1]/shell[1]").hierarchyHeaderWidth = 480
            self._session.findById("wnd[0]/usr/subAREA06:SAPLCCM21:0101/tabsTABSTRIP/tabpAA_TAB1/ssubSUB_TAB:SAPLCCM21:0105/ssubCCM21_CUST_SUB:SAPLCCM1:0501/subSEARCH_DISPLAY:SAPLCRM_CIC_BP_SUB:6006/cntlHTML_CONTROL_BPSEARCH/shellcont/shell").sapEvent("","BP1_PARTNER=&BP1_MCNAME1=&BP1_MCNAME2=&C_TAXNUM=&BP2_PARTNER=&BP2_MCNAME1=&BP2_MCNAME2=&TIME_ZONE=&A_VKONT=&SEARCHSTRING=&ADRTYP=BP1&STREET=&HOUSE_NUM1=&CITY1=&POST_CODE1=&COUNTRY=&TELEPHONE=&TEL_EXT=&FAX=&FAX_EXT=&EMAIL=","sapevent:ISU_FIND")
            self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2").select()
        except:                 
            pass
        
        try:
            self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-I_ANLAGE").text = "33052727"
            self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-I_ANLAGE").setFocus()
            self._session.findById("wnd[1]/usr/tabsSEARCHFIELDS/tabpTAB2/ssubSUB2:SAPLEFND:0112/ctxtEFINDD-I_ANLAGE").caretPosition = 8
            self._session.findById("wnd[1]").sendVKey(0)
        except:
             pass
        
        try:
            self._session.findById("wnd[0]/tbar[0]/btn[12]").press()
            self._session.findById("wnd[0]/tbar[0]/btn[12]").press()
            self._session.findById("wnd[0]/usr/subAREA06:SAPLCCM21:0101/tabsTABSTRIP/tabpAA_TAB2").select()
            self._session.findById("wnd[0]/usr/subAREA06:SAPLCCM21:0101/tabsTABSTRIP/tabpAA_TAB2/ssubSUB_TAB:SAPLCCM21:0102/cntlCCM21_CUST_CONT/shellcont/shell/shellcont[1]/shell").sapEvent("","","sapevent:?OBJTYPE=INSTLN&OBJKEY=0033052727&LOGSYS=EP1CLNT500")
            self._session.findById("wnd[1]/usr/lbl[1,1]").caretPosition = 7
            self._session.findById("wnd[1]").sendVKey(2)
        except:
             pass
        
        try:
            self._session.findById("wnd[0]/tbar[0]/btn[3]").press()
            self._session.findById("wnd[0]/usr/subAREA04:SAPLCRM_CIC_SLIM_ACTION_BOX:0100/cntlCRMCICSLIMABCONTAINER/shellcont/shell").pressContextButton("0009")
            self._session.findById("wnd[0]/usr/subAREA04:SAPLCRM_CIC_SLIM_ACTION_BOX:0100/cntlCRMCICSLIMABCONTAINER/shellcont/shell").selectContextMenuItem("CAAT")
            self._session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").text = "33052727"
            self._session.findById("wnd[1]/usr/sub:SAPLSPO4:0300/ctxtSVALD-VALUE[0,21]").caretPosition = 8
            self._session.findById("wnd[1]/tbar[0]/btn[0]").press()
            self._session.findById("wnd[1]/usr/btnBT_CONT").press()
            self._session.findById("wnd[1]/usr/btnBUTTON_1").press()
            self._session.findById("wnd[1]/tbar[0]/btn[5]").press()
        except:
            pass

        return "Funcionalidade em desenvolvimento"

    def close(self) -> None:
        self._connection.CloseSession('ses[0]')

    def back(self) -> None:
        self._session.findById("wnd[0]/tbar[0]/btn[3]").press()
        self._session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()
