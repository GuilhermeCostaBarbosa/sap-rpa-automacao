import pandas as pd
import win32com.client
from time import time, sleep
from tqdm import tqdm

df = pd.read_excel('tabela_materiais.xlsx')

session = None  

def connect_sap():
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        raise Exception("SAP GUI não encontrado. Abra e faça o login.")
    
    application = SapGuiAuto.GetScriptingEngine

    for connection in application.Connections:
        for sess in connection.Sessions:
            try:
                sess.findById("wnd[0]")
                print("Conectado ao SAP com sucesso.")
                return sess
            except Exception:
                pass
    raise Exception("Nenhuma sessão SAP ativa encontrada. Abra uma sessão e faça o login.")

def get_sap_session():
    global session
    if session is None:
        session = connect_sap()
    return session

def wait_for_element(element_id, timeout=15):
    sess = get_sap_session()
    start_time = time()
    while time() - start_time < timeout:
        try:
            element = sess.findById(element_id)
            return element
        except Exception:
            sleep(0.5)
    raise Exception(f"Elemento SAP não encontrado: {element_id}")

    # Acessar a transação LOF
sess = get_sap_session()

def update_lof(planilha):
    # Acessar a transação LOF
    wait_for_element("wnd[0]/tbar[0]/okcd").text = "ME01"
    sess.findById("wnd[0]").sendVKey(0)
    for itens in tqdm(planilha.index):
        material = str(df.loc[itens, 'Cod_Sap'])
        centro = str(df.loc[itens, 'centro'])
        in_vigencia = df.loc[itens, 'in_vig']
        fim_vigencia = df.loc[itens, 'fim_vig']
        fornecedor = str(df.loc[itens, 'Fornecedor'])
        org_compras = str(df.loc[itens, 'OrgC'])
        contrato = str(df.loc[itens, 'Contrato'])
        item = str(df.loc[itens, 'Item'])


    # Preencher os campos necessários
        sess.findById("wnd[0]/usr/ctxtEORD-MATNR").text = material
        sess.findById("wnd[0]/usr/ctxtEORD-WERKS").text = centro

        sess.findById("wnd[0]").sendVKey(0)

        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").selected = True
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,0]").text = in_vigencia
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,0]").text = fim_vigencia
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text = fornecedor
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,0]").text = org_compras
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EBELN[6,0]").text = contrato
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EBELP[7,0]").text = item
        sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").setFocus()
        sess.findById("wnd[0]/tbar[0]/btn[11]").press()
        sess.findById("wnd[0]/tbar[0]/btn[11]").press()

def flags():
    sess = get_sap_session()
    for itens in tqdm(df.index):
        material = str(df.loc[itens, 'Cod_Sap'])
        centro = int(df.loc[itens, 'centro'])
        sess.findById("wnd[0]").maximize()
        sess.findById("wnd[0]/tbar[0]/okcd").text = "MM02"
        sess.findById("wnd[0]").sendVKey(0)
        sess.findById("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
        sess.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 7
        sess.findById("wnd[0]").sendVKey(0)
        sess.findById("wnd[1]/tbar[0]/btn[0]").press()
        wait_for_element("wnd[1]/usr/ctxtRMMG1-WERKS").text = centro
        sess.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
        sess.findById("wnd[1]/tbar[0]/btn[0]").press()
        sess.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").selected = True
        sess.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").setFocus()
        sess.findById("wnd[0]/tbar[0]/btn[11]").press()
        

update_lof(df)
flags()
