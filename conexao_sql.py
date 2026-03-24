import pyodbc
import pandas as pd

# 1. Definindo os parâmetros de conexão
# Usamos Trusted_Connection=yes para Autenticação do Windows
dados_conexao = (
    "Driver={SQL Server};"
    "Server=R_AXELL12\SQLEXPRESS;"
    "Database=ProjetoVendas;"  # <--- COLOQUE O NOME DO SEU BANCO AQUI
    "Trusted_Connection=yes;"
)

try:
    # 2. Tentando estabelecer a conexão
    conexao = pyodbc.connect(dados_conexao)
    print("Conexão com SQL Server realizada com sucesso!")

    # 3. Criando a Query (Pergunta) para o Banco
    # Substitua pelo nome da tabela que você importou do Excel
    query = "SELECT * FROM VENDAS"

    # 4. Carregando os dados em um DataFrame do Pandas
    df = pd.read_sql(query, conexao)

    # 5. Exibindo as 5 primeiras linhas para validar
    print("\n--- Primeiras linhas dos seus dados ---")
    print(df.head())

    # 6. Uma análise rápida (Estatística Descritiva)
    print("\n--- Resumo Estatístico ---")
    print(df.describe())

    # Fechar a conexão após o uso
    conexao.close()

except Exception as e:
    print(f"Ocorreu um erro: {e}")
