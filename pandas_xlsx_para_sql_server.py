# PyCharm 2021.2.3
# Importa as bibliotecas
# "pandas" é utilizada para a manipulação e análise de dados
# "pyodbc" permite a conexão com o banco de dados, neste caso, SQL Server
import pandas as pd
import pyodbc

# Define as variáveis de conexão com o banco de dados SQL Server
v_server = '(Local)'
v_database = 'database'
v_username = 'username'
v_password = 'password'
v_conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+v_server+';DATABASE='+v_database+';UID='+v_username+';PWD='+ v_password)
v_cursor = v_conn.cursor()

# Faz a leitura do arquivo (.xlsx)
# Renomeia o titulo das colunas
v_tabela = pd.read_excel(r'...\Aracaju.xlsx', sheet_name='Planilha1')
v_data = v_tabela.rename(columns={'Cidade': 'CIDADE',
                            'Data': 'DATA_VENDA',
                            'Vendas': 'VALOR_VENDA',
                            'LojaID': 'LOJA',
                            'Qtde': 'QUANTIDADE'})

# Elimina todos os registros da tabela antes da inserir os novos registros
truncsql = 'TRUNCATE TABLE TABELA'
v_cursor.execute(truncsql)

# Alguns exemplos para visualização dos dados do arquivo
#print(v_data.head())
#print(v_data.columns)
#print(v_data.tail(10))
#print(v_data.describe())
#print(v_data["QUANTIDADE"].sum(), v_data["VALOR VENDA"].sum())

# Define a variável com o comando de inserção dos novos registros na tabela
v_query = """INSERT INTO TABELA (CIDADE, DATA_VENDA, VALOR_VENDA, LOJA, QUANTIDADE) VALUES (?, ?, ?, ?, ?)"""

# Executa a estrutura de repetição fazendo a leitura dos registros do arquivo e atribui as variáveis
for i in v_data.index:
    CIDADE = v_data['CIDADE'][i]
    DATA_VENDA = v_data['DATA_VENDA'][i]
    VALOR_VENDA = v_data['VALOR_VENDA'][i]
    LOJA = str(v_data['LOJA'][i])
    QUANTIDADE = str(v_data['QUANTIDADE'][i])

    v_values = (CIDADE, DATA_VENDA, VALOR_VENDA, LOJA, QUANTIDADE)
    # Executa o comando de inserção dos novos registros
    v_cursor.execute(v_query, v_values)
# Confirma a gravação dos registros na base de dados
v_conn.commit()
