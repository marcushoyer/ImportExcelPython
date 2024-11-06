import pandas as pd
from sqlalchemy import create_engine
import concurrent.futures
import os
import glob
#import config
from config import DevelopmentConfig, ProductionConfig

# Seleciona a configuração com base no ambiente
if os.getenv('FLASK_ENV') == 'production':
    config = ProductionConfig()
    print("Conectando em PRD.")
else:
    print("Conectando em DSV.")
    config = DevelopmentConfig()

# Cria a conexão com o banco de dados usando `SQLALCHEMY_DATABASE_URI` da configuração
engine = create_engine(config.SQLALCHEMY_DATABASE_URI)

def process_sheet(sheet_name, df):
    """Função para salvar a aba em uma tabela do banco de dados"""
    # Convertendo o nome da aba para minúsculas para evitar conflitos de nome
    table_name = sheet_name.lower()

    # Salvando o DataFrame no PostgreSQL usando chunksize para controle de memória
    df.to_sql(table_name, engine, if_exists='replace', index=False, method='multi', chunksize=1000)
    print(f"Tabela '{table_name}' criada e dados inseridos.")

def read_excel_and_import(file_path):
    """Função principal para ler o arquivo Excel e processar cada aba"""
    # Carregando o arquivo Excel
    try:
        excel_data = pd.ExcelFile(file_path)
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return

    # Usando ThreadPoolExecutor para processar as abas em paralelo
    with concurrent.futures.ThreadPoolExecutor() as executor:
        futures = []
        for sheet_name in excel_data.sheet_names:
            # Lendo cada aba
            df = excel_data.parse(sheet_name)
            # Submetendo a tarefa para processar cada aba individualmente
            futures.append(executor.submit(process_sheet, sheet_name, df))

        # Aguardando todas as tarefas serem completadas
        for future in concurrent.futures.as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"Erro ao processar uma aba: {e}")

if __name__ == "__main__":
    # Definindo o caminho para o arquivo Excel
    file_path = '/media/marcus/WD2TB/Projetos/ImportExcel/'

    # Listar todos os arquivos .xlsx na pasta
    for arquivo in glob.glob(f"{file_path}/*.xlsx"):
        # Checando se o arquivo existe
        if os.path.isfile(arquivo):
            read_excel_and_import(arquivo)
        else:
            print(f"O arquivo '{arquivo}' não foi encontrado.")
