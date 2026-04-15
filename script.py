import pandas as pd
from time import time, sleep
from tqdm import tqdm
import logging
import sys

# Configuração de Logs
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('script_sap.log', mode='a', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)


# Validação de dados
def validar_planilha(caminho_arquivo):
    logging.info(f"Lendo arquivo: {caminho_arquivo}")
    # Verificando se arquivo existe
    try:
        df = pd.read_excel(caminho_arquivo)
    except FileNotFoundError:
        logging.critical(f"Arquivo não encontrado: {caminho_arquivo}")
        raise
    
    # verificando de arquivo está vázio
    if df.empty:
        logging.error(f"A planilha carregada está vazia")
        raise ValueError("A planilha carregada está vazia")

    # Validando se o arquivo possui todas as colunas necessárias para automação
    col_necessarias = ['Cod_Sap', 'centro', 'in_vig', 'fim_vig', 'Fornecedor', 'OrgC', 'Contrato', 'Item']
    for col in col_necessarias:
        # Caso faltar colunas, retornar log de erro e informar colunas ausente
        if col not in df.columns:
            logging.error(f"Coluna necessária ausente: {col}")
            raise KeyError(f"Erro: ausencia da coluna {col} na planilha")
    
    # Caso validado, retornar df
    logging.info(f"Planilha validada com sucesso, número total de linhas é de {len(df)}")
    return df

session = None  

# Inicando conexão com o SAP
def connect_sap():
    import win32com.client
    logging.info("Tentando conecetar co SAP GUI...")
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
    except Exception:
        logging.critical("SAP GUI não encontrado.")
        raise Exception("SAP GUI não encontrado. Abra e faça o login.")
    
    application = SapGuiAuto.GetScriptingEngine

    for connection in application.Connections:
        for sess in connection.Sessions:
            try:
                sess.findById("wnd[0]")
                print("Conectado ao SAP com sucesso.")
                logging.info("Conectado ao SAP com sucesso.")
                return sess
            except Exception:
                pass
    logging.error("Nenhuma sessão SAP ativa foi encontrada")
    raise Exception("Nenhuma sessão SAP ativa encontrada. Abra uma sessão e faça o login.")

# Capturando sessão SAP
def get_sap_session():
    global session
    if session is None:
        session = connect_sap()
    return session

# Espera para localização dos elementos SAP pelo element_id
def wait_for_element(element_id, timeout=15):
    sess = get_sap_session()
    start_time = time()

    # Loop de espera até que atinja timeout
    while time() - start_time < timeout:
        try:
            element = sess.findById(element_id)
            return element
        except Exception:
            sleep(0.5)
    # Caso não encontre o elemento, informa o id e o timeout
    logging.error(f'Timeout: o elemento {element_id} não carregou após {timeout}')
    raise Exception(f"Elemento SAP não encontrado: {element_id}")

# Atualiza o registro LOF (ME01) conforme DF
def update_lof(sess, df):
    # Acessar a transação LOF
    logging.info('Iniciando atualização LOF (ME01)...')
    wait_for_element("wnd[0]/tbar[0]/okcd").text = "ME01"
    sess.findById("wnd[0]").sendVKey(0)

    # Loop para a lof de cada item em DF
    for itens in tqdm(df.index, desc='Atualizando LOF (ME01)'):
        try:
            # Atribuindo valor de cada item nas variaveis para cada campo
            material = str(df.loc[itens, 'Cod_Sap'])
            centro = str(df.loc[itens, 'centro'])
            in_vigencia = df.loc[itens, 'in_vig']
            fim_vigencia = df.loc[itens, 'fim_vig']
            fornecedor = str(df.loc[itens, 'Fornecedor'])
            org_compras = str(df.loc[itens, 'OrgC'])
            contrato = str(df.loc[itens, 'Contrato'])
            item = str(df.loc[itens, 'Item'])


            # Preencher os campos necessários para cadastrar LOF do item
            wait_for_element("wnd[0]/usr/ctxtEORD-MATNR").text = material
            sess.findById("wnd[0]/usr/ctxtEORD-WERKS").text = centro

            # Acessar tela de preenchimento de dados
            sess.findById("wnd[0]").sendVKey(0)

            # preencher cada campo e salvar modificações
            wait_for_element("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").selected = True
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-VDATU[0,0]").text = in_vigencia
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-BDATU[1,0]").text = fim_vigencia
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-LIFNR[2,0]").text = fornecedor
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EKORG[3,0]").text = org_compras
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EBELN[6,0]").text = contrato
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/ctxtEORD-EBELP[7,0]").text = item
            sess.findById("wnd[0]/usr/tblSAPLMEORTC_0205/chkRM06W-FESKZ[8,0]").setFocus()
            sess.findById("wnd[0]/tbar[0]/btn[11]").press()
            sess.findById("wnd[0]/tbar[0]/btn[11]").press()

        except Exception as e:
            # Caso algum item der erro, informar material em log e prosseguir com o restante
            logging.warning(f"Falha ao processar material {material} na ME01: {e}")
            continue
    
    # Sair da transação ao finalizar o loop
    sess.findById("wnd[0]/tbar[0]/btn[15]").press()

#  Ativa flag para emissão de pedido automático após vincular lof
def flags(sess, df):
    logging.info('Iniciando automação de flags (MM02)...')

    # Acessar a transação MM02
    sess.findById("wnd[0]").maximize()
    wait_for_element("wnd[0]/tbar[0]/okcd").text = "MM02"
    sess.findById("wnd[0]").sendVKey(0)

    # Loop para atualizar flag de cada item em DF
    for itens in tqdm(df.index, desc='Atualizando flags (MM02)'):
        try:
            # atribuindo valores a variaveis
            material = str(df.loc[itens, 'Cod_Sap'])
            centro = int(df.loc[itens, 'centro'])

            # Preenchendo campos necessarios
            wait_for_element("wnd[0]/usr/ctxtRMMG1-MATNR").text = material
            sess.findById("wnd[0]/usr/ctxtRMMG1-MATNR").caretPosition = 7
            sess.findById("wnd[0]").sendVKey(0)
            wait_for_element("wnd[1]/tbar[0]/btn[0]").press()
            wait_for_element("wnd[1]/usr/ctxtRMMG1-WERKS").text = centro
            sess.findById("wnd[1]/usr/ctxtRMMG1-WERKS").caretPosition = 4
            sess.findById("wnd[1]/tbar[0]/btn[0]").press()
            
            # Atualizar flag de pedido automatico
            wait_for_element("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").selected = True
            sess.findById("wnd[0]/usr/tabsTABSPR1/tabpSP10/ssubTABFRA1:SAPLMGMM:2000/subSUB2:SAPLMGD1:2301/chkMARC-KAUTB").setFocus()

            # salvar alteracao
            sess.findById("wnd[0]/tbar[0]/btn[11]").press()
        except Exception as e:
            logging.warning(f'Falhar ao atualizar o material {material} na MM02: {e}')
            continue
    
    # Sair da transação ao finalizar o loop
    sess.findById("wnd[0]/tbar[0]/btn[15]").press()

if __name__ == '__main__':
    NOME_ARQUIVO = 'tabela_materiais.xlsx'

    try:
        logging.info('Iniciando automação SAP')
        df_dados = validar_planilha(NOME_ARQUIVO)

        sessao_sap = get_sap_session()

        update_lof(sessao_sap, df_dados)
        flags(sessao_sap, df_dados)

        logging.info('Automação realizada com sucesso!')
    except Exception as e:
        logging.critical(f'Automacao interrompida devido ao erro: {e}')
