import pandas as pd
import os

def get_latest_xlsx_file(target_dir):
    """
    Recupera o arquivo .xlsx mais recente adicionado no diretório de destino.

    Parâmetros:
        target_dir (Path): Caminho para a pasta de destino.

    Retorna:
        Path: Caminho para o arquivo .xlsx mais recente.
        None: Se nenhum arquivo .xlsx for encontrado.
    """
    xlsx_files = list(target_dir.glob('*.xlsx'))
    if not xlsx_files:
        print("No .xlsx files found in the CrawledData folder.")
        return None

    # Classificando os arquivos pela data de criação em ordem decrescente e retornando o primeiro.
    latest_file = max(xlsx_files, key=lambda f: f.stat().st_ctime)
    print(f"Latest file identified: {latest_file.name}")
    return latest_file


def load_xlsx_to_dataframe(file_path):
    """
    Passa um arquivo .xlsx para um DataFrame do pandas.

    Parâmetros:
        file_path (Path): Caminho para o arquivo .xlsx.

    Retorna:
        pandas.DataFrame: DataFrame contendo os dados do Excel.
    """
    try:
        df = pd.read_excel(file_path, engine='openpyxl')  # Specify engine if necessary
        print(f"Loaded {file_path.name} into DataFrame.")
        return df
    except Exception as e:
        print(f"Error loading {file_path.name}: {e}")
        return None


def processar_relatorio(df_relatorio, df_municipios):
    """
    Limpa e transforma a coluna 'Valor de Repasse' em float,
    adiciona a coluna 'id' de df_municipios como foreign key
    e renomeia a coluna para 'Municipio_id'.

    Parâmetros:
        df_relatorio (pd.DataFrame): DataFrame contendo o relatório com coluna 'Valor de Repasse' e 'Município'.
        df_municipios (pd.DataFrame): DataFrame contendo os dados dos municípios com coluna 'id' e 'nome_mun'.

    Retorna:
        pd.DataFrame: DataFrame processado com a coluna 'Valor de Repasse' em float e a fk 'Municipio_id'.
    """

    # Converter a coluna "Valor de Repasse" para float
    df_relatorio['Valor de Repasse'] = df_relatorio['Valor de Repasse'].replace(r'[$,]', '', regex=True).astype(float)

    # Fazer o merge para adicionar o 'id' de df_municipios como foreign key
    df_relatorio = df_relatorio.merge(df_municipios[['id', 'nome_mun']], left_on='Município', right_on='nome_mun',
                                      how='left')

    # Renomear a coluna 'id' e remover a coluna auxiliar 'nome_mun'
    df_relatorio = df_relatorio.rename(columns={'id': 'Municipio_id'})
    df_relatorio.drop(['nome_mun'], axis=1, inplace=True)

    return df_relatorio


def save_dataset_to_xlsx(dataset, file_path='Output/relatorio_final.xlsx'):
    """
    Salva o conjunto de dados fornecido em um arquivo Excel chamado 'relatorio_final.xlsx'.

    Parâmetros:
        dataset (pd.DataFrame): O conjunto de dados a ser salvo.
        file_path (str): O caminho onde o arquivo Excel será salvo.
    """
    try:
        # Confirmar que o diretório existe
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        # Salvar o Dataframe como Excel
        dataset.to_excel(file_path, index=False, engine='openpyxl')
        print(f"Dataset successfully saved to {file_path}")
    except Exception as e:
        print(f"Failed to save dataset to Excel: {e}")