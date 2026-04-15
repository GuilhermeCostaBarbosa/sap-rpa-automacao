# 🤖 SAP RPA: Automação de Suprimentos & Engenharia de Confiabilidade

[![Python Version](https://img.shields.io/badge/python-3.12-blue.svg)](https://python.org)
[![Pytest](https://img.shields.io/badge/pytest-passing-success.svg)](https://pytest.org)
[![DevOps Mindset](https://img.shields.io/badge/SRE-Toil_Reduction-orange.svg)]()

Um robô (RPA) desenvolvido em Python para automatizar a atualização em massa da Lista de Opções de Fornecimento (LOF - ME01) e Flags de Pedido Automático (MM02) no SAP GUI.

## 🎯 O Problema de Negócio (Redução de Toil)
A rotina de atualização de contratos e listas de fornecimento de materiais em ambientes hospitalares exige a digitação manual de milhares de linhas e transições de tela no SAP. Isso gerava horas de **Toil** (trabalho braçal, repetitivo e sem valor agregado), além de abrir margem para erros humanos de digitação, impactando a cadeia de suprimentos.

## 💡 A Solução
Desenvolvi um script parametrizado que extrai dados de planilhas locais e interage nativamente com a API `win32com` do Windows para operar a interface do SAP em alta velocidade e com validações de segurança.

### 🏗️ Destaques da Arquitetura (Práticas DevOps)
Este projeto foi construído com foco em resiliência e observabilidade corporativa:

* **Resiliência:** Implementação de uma função de espera inteligente (`wait_for_element`) nas transições de tela. O robô aguarda o carregamento dinâmico dos componentes do SAP em vez de usar pausas fixas, evitando quebras de script causadas por latência na rede.
* **Observabilidade:** Remoção de saídas de console comuns (`prints`) por uma biblioteca de `logging` estruturada. O sistema gera um arquivo `script_sap.log` com data, hora e níveis de severidade (`INFO`, `WARNING`, `CRITICAL`), essencial para auditoria e troubleshooting.
* **Fail-Fast & CI Ready:** Implementação de testes unitários com `pytest` para a lógica de higienização de dados (Pandas). O sistema isola a dependência do Windows para rodar em ambientes Linux (CI/CD) e aplica o conceito de "Falha Rápida": aborta a operação caso detecte planilhas vazias ou colunas ausentes antes mesmo de tentar abrir o SAP.

## 🚀 Tecnologias Utilizadas
* **Python 3.12**
* **Pandas:** Sanitização e validação de DataFrames.
* **Pywin32:** Integração COM (Component Object Model) com o SAP GUI.
* **Pytest:** Esteira de testes automatizados e simulação de falhas (Negative Testing).
* **Tqdm:** Observabilidade local com barra de progresso no terminal.

## 📂 Estrutura do Projeto

├── .gitignore                          
├── requirements.txt         
├── script.py                
├── test_script.py          
└── tabela_materiais.xlsx  

## ⚙️ Como Executar

**1. Instale as dependências:**
```bash
pip install -r requirements.txt
```

**2. Para rodar a Esteira de Testes (Validação de Dados):**

**Os testes** podem seres executados em ambientes **Linux ou Windows**

```bash
pytest test_script.py
```
**3. Rodar Automação no SAP**`

**Para rodar a Automação** requer ambiente **Windows com SAP GUI logado.**
```bash
python script.py
```

---

Desenvolvido com foco na eliminação de trabalho repetitivo e melhoria do processo.

---

## Autor
Guilherme Costa Barbosa

LinkedIn: https://www.linkedin.com/in/guilherme-costa-barbosa-345178261/
