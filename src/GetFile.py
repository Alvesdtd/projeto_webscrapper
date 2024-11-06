import shutil
from pathlib import Path
from datetime import datetime, timedelta
import sys

def move_recent_xlsx(download_dir, target_dir, time_window_minutes=1):
    """
    Move arquivos .xlsx criados nos últimos minutos de `download_dir` para `target_dir`.

    Parâmetros:
        download_dir (Path): Caminho para a pasta de Downloads.
        target_dir (Path): Caminho para a pasta de destino.
        time_window_minutes (int): Janela de tempo em minutos para considerar arquivos recentes.
    """
    # Confirma que o diretório existe
    target_dir.mkdir(parents=True, exist_ok=True)

    # Setta o "Current time"
    now = datetime.now()
    time_threshold = now - timedelta(minutes=time_window_minutes)

    # Itera por todos os arquivos .xlsx no diretório alvo
    for file_path in download_dir.glob('*.xlsx'):
        try:
            # Busca quando o arquivo foi criado
            file_stat = file_path.stat()
            creation_time = datetime.fromtimestamp(file_stat.st_ctime)

            # Checka se o arquivo foi criado na janela de tempo estipulada
            if creation_time >= time_threshold:
                # Define o destino
                destination = target_dir / file_path.name

                # Move o arquivo
                shutil.move(str(file_path), str(destination))
                print(f"Moved: {file_path.name} to {destination}")
        except Exception as e:
            print(f"Error processing file {file_path.name}: {e}")

def get_default_download_folder():
    """
    Retorna o diretório padrão de Downloads com base no sistema operacional.
    """
    home = Path.home()
    if sys.platform.startswith('win'):
        return home / 'Downloads'
    elif sys.platform.startswith('darwin'):
        return home / 'Downloads'
    else:
        # Linux e outros
        return home / 'Downloads'