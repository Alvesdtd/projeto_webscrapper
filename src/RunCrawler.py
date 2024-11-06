from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.chrome.options import Options


def download_emendas(
    url='https://paineis.cidadania.gov.br/public/extensions/RFF/emendas.html',
    button_id='btn_export',
    headless=True,
    window_size=(1920, 1080),
    page_load_timeout=30,
    download_wait_time=15
):
    """
    Faz o download do arquivo Excel de emendas a partir da URL especificada.

    Parâmetros:
        url (str): A URL para navegar.
        button_id (str): O ID do botão de exportação/download.
        headless (bool): Define se o Chrome será executado no modo headless.
        window_size (tuple): O tamanho da janela para o navegador.
        page_load_timeout (int): Tempo máximo para esperar o carregamento da página.
        download_wait_time (int): Tempo de espera após clicar em download.
    """
    chrome_options = Options()
    if headless:
        chrome_options.add_argument('--headless')  # Não abrir janela gráfica
    chrome_options.add_argument(f'--window-size={window_size[0]},{window_size[1]}')  # Configurar tamanho da janela

    # Inicializar o WebDriver
    driver = webdriver.Chrome(options=chrome_options)

    try:
        # Navegar para a página
        driver.get(url)
        print("Inicializando o Crawler")

        # Esperar a página carregar
        WebDriverWait(driver, page_load_timeout).until(
            EC.presence_of_element_located((By.TAG_NAME, 'body'))
        )

        # Esperar o botão de download carregar e ser clicável
        wait = WebDriverWait(driver, 30)
        download_button = wait.until(EC.element_to_be_clickable((By.ID, button_id)))

        # Scroll na página até o botão
        driver.execute_script("arguments[0].scrollIntoView(true);", download_button)
        time.sleep(2)  # Tempo para assegurar que o scroll chegou até o botão

        # Clicar no botão
        try:
            download_button.click()
            print("Realizando o Download do Arquivo")
        except Exception:
            # Se falhar no HTML, tentar por JS
            driver.execute_script("arguments[0].click();", download_button)

        # Esperar o download terminar
        time.sleep(download_wait_time)

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Finalizar o Browser
        driver.quit()
        print("Finalizando o Crawler")