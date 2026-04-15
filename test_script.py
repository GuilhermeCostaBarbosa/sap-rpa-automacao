import pytest
import pandas as pd
from script import validar_planilha

# Teste para evitar planilhas vazias
def test_planilha_vazia(tmp_path):
    df_vazio = pd.DataFrame()
    arquivo_falso = tmp_path / 'vazia.xlsx'
    df_vazio.to_excel(arquivo_falso, index=False)

    with pytest.raises(ValueError, match='A planilha carregada está vazia'):
        validar_planilha(arquivo_falso)

def test_sem_col_sap(tmp_path):
    dados_errados = {
        'Codigo_errado': ['123456'],
        'centro': [1234],
        'in_vig': ['01.01.2026'],
        'fim_vig': ['31.12.2026'],
        'Fornecedor': ['654321'],
        'OrgC': ['OC02'],
        'Contrato': ['4600001234'],
        'Item': ['10']
    }

    df_falso = pd.DataFrame(data=dados_errados)
    arquivo_falso = tmp_path / 'errada.xlsx'
    df_falso.to_excel(arquivo_falso, index=False)

    with pytest.raises(KeyError, match='Erro: ausencia da coluna Cod_Sap na planilha'):
        validar_planilha(arquivo_falso)

def test_planilha_valida(tmp_path):
    dados_corretos = {
        'Cod_Sap': ['123456'],
        'centro': [1234],
        'in_vig': ['01.01.2026'],
        'fim_vig': ['31.12.2026'],
        'Fornecedor': ['654321'],
        'OrgC': ['OC02'],
        'Contrato': ['4600001234'],
        'Item': ['10']
    }

    df_certos = pd.DataFrame(data=dados_corretos)
    arquivo_certo = tmp_path / 'certa.xlsx'
    df_certos.to_excel(arquivo_certo, index=False)

    resultado = validar_planilha(arquivo_certo)
    assert not resultado.empty