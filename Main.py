from pathlib import Path
from src.GetFile import get_default_download_folder, move_recent_xlsx
from src.RunCrawler import download_emendas
from src.DataTreatment import get_latest_xlsx_file, load_xlsx_to_dataframe, processar_relatorio, save_dataset_to_xlsx

def main():
    # Chamando o Crawler
    download_emendas()

    # Definindo o diretório de download
    download_dir = get_default_download_folder()
    target_dir = Path(__file__).parent.resolve() / 'CrawledData'

    print(f"Download Directory: {download_dir}")
    print(f"Target Directory: {target_dir}")

    # Movendo arquivos .xlsx recentes
    move_recent_xlsx(download_dir, target_dir, time_window_minutes=1)
    latest_file = get_latest_xlsx_file(target_dir)

    # Transformando o .xlsx do Relatório Detalhado em um dataframe
    df_relatorio = load_xlsx_to_dataframe(latest_file)

    # Transformando o .xlsx da Tabela Municipios em um dataframe
    file_path = Path('Data/TABELA_MUNICIPIOS.xlsx')
    df_municipios = load_xlsx_to_dataframe(file_path)

    # Montando o Relatório Final
    relatorio_final = processar_relatorio(df_relatorio, df_municipios)

    # Transformando o dataset final em arquivo .xlsx (Output em excel)
    save_dataset_to_xlsx(relatorio_final)

if __name__ == "__main__":
    main()