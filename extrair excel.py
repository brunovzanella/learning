import pyodbc
import pandas as pd

server = 'INSERIR IP DO BANCO'
database = 'INSERIR BASE DE DADOS'
username = 'NOME DE USUÁRIO'
password = 'SENHA'

conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};'
                      f'SERVER={server};'
                      f'DATABASE={database};'
                      f'UID={username};'
                      f'PWD={password}'
)
query = 'INSERIR QUERY'

df = pd.read_sql(query, conn)

conn.close()

df.to_excel(r'C:\Users\dados_banco.xlsx', index=False)

print('Planilha salva com sucesso')
