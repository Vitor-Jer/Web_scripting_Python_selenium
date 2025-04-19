from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager
import time
import os
import logging
import zipfile

# Configuração básica do logging
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('automation.log'),  # Grava logs em um arquivo
        logging.StreamHandler()  # Exibe logs no console
    ]
)

# Função para realizar login com E-mail
def email_login(driver, email, senha):
    try:
        logging.info("Iniciando login com E-mail.")
        
        # Clica no botão "Sign in with Email"
        login_button = driver.find_element(By.XPATH, "//span[contains(text(),'Sign in with Email')]")
        login_button.click()
        time.sleep(3)

        # Preenche o e-mail
        email_field = driver.find_element(By.NAME, "email")
        email_field.send_keys(email)

        # Clica no botão de continuar
        continue_button = driver.find_element(By.XPATH, "//button[@type='submit']")
        continue_button.click()
        time.sleep(3)

        # Preenche a senha
        senha_field = driver.find_element(By.NAME, "password")
        senha_field.send_keys(senha)

        # Clica no botão de login
        login_button = driver.find_element(By.XPATH, "//button[@type='submit']")
        login_button.click()

        # Aguarda alguns segundos para o login ser processado
        time.sleep(5)

        # Verifica se o login foi bem-sucedido
        if "login" not in driver.current_url:
            logging.info("Login realizado com sucesso!")
        else:
            raise Exception("Falha no login. Verifique suas credenciais.")
    except NoSuchElementException as e:
        logging.error(f"Erro ao tentar fazer login com E-mail: {e}")
    except Exception as e:
        logging.error(f"Erro inesperado durante o login: {e}")

# Função para extrair o arquivo ZIP
def extract_zip(download_dir, zip_filename):
    zip_path = os.path.join(download_dir, zip_filename)
    extract_to = os.path.join(download_dir, 'extracted_files')
    
    try:
        logging.info(f"Extraindo arquivo ZIP: {zip_path}")
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(extract_to)  # Extrai para a pasta 'extracted_files'
        logging.info(f"Extração concluída. Arquivos extraídos para: {extract_to}")
    except FileNotFoundError:
        logging.error(f"Arquivo {zip_filename} não encontrado.")
    except zipfile.BadZipFile:
        logging.error(f"Erro ao extrair {zip_filename}: Arquivo ZIP corrompido.")
    except Exception as e:
        logging.error(f"Erro inesperado ao extrair o arquivo ZIP: {e}")

# Função principal de automação
def automate_download(url, email, senha):
    # Configuração do WebDriver (usando WebDriver Manager)
    chrome_options = Options()

    # Configura o local para onde os arquivos serão baixados (pasta de downloads do sistema)
    download_dir = os.path.join(os.getenv('USERPROFILE'), 'Downloads')  # Pasta de downloads do sistema no Windows
    prefs = {
        "download.default_directory": download_dir,  # Define a pasta de downloads do sistema
        "download.prompt_for_download": False,  # Baixa automaticamente sem pedir confirmação
        "safebrowsing.enabled": True  # Evita avisos de downloads inseguros
    }
    chrome_options.add_experimental_option("prefs", prefs)
    service = Service(ChromeDriverManager().install())

    # Inicializa o WebDriver
    driver = webdriver.Chrome(service=service, options=chrome_options)

    try:
        logging.info(f"Acessando a URL: {url}")
        # Navega para a página de download
        driver.get(url)
        time.sleep(3)  # Espera carregar a página

        # Localiza e clica no botão de download
        try:
            download_btn = driver.find_element(By.XPATH, "//span[contains(text(),'Download')]")
            download_btn.click()
            logging.info("Botão de download clicado. Aguardando login...")

            # Aguarda a página de login carregar
            time.sleep(5)

            # Faz o login com E-mail
            email_login(driver, email, senha)

            # Tenta clicar no botão de download novamente após o login, se necessário
            try:
                download_btn = driver.find_element(By.XPATH, "//span[contains(text(),'Download')]")
                download_btn.click()
                logging.info("Download iniciado com sucesso.")
                time.sleep(10)  # Espera o download iniciar e ser concluído

                # Verifica se o arquivo foi baixado
                zip_filename = 'archive.zip'  # Nome do arquivo ZIP baixado
                zip_file_path = os.path.join(download_dir, zip_filename)
                if os.path.exists(zip_file_path):
                    logging.info(f"Arquivo baixado: {zip_file_path}")
                    
                    # Extrai o arquivo ZIP
                    extract_zip(download_dir, zip_filename)
                else:
                    logging.error("Arquivo ZIP não encontrado no diretório de downloads.")

            except NoSuchElementException:
                logging.error("Botão de download não encontrado após login.")
        except NoSuchElementException:
            logging.error("Botão de download não encontrado.")
        except TimeoutException:
            logging.error("Tempo limite excedido ao tentar acessar o site.")
        except Exception as e:
            logging.error(f"Erro inesperado: {e}")

    except Exception as e:
        logging.error(f"Erro ao acessar o site: {e}")
    finally:
        driver.quit()
        logging.info("WebDriver encerrado.")

if __name__ == "__main__":
    # URL do site para baixar o arquivo
    site_url = 'https://www.kaggle.com/datasets/thomassimeo/contas-senior-neobpo?select=DESPESAS.xlsx'

    # Credenciais de login (e-mail e senha)
    user_email = input('E-mail: ')  # Substitua pelo seu e-mail
    user_password = input('Senha: ')  # Substitua pela sua senha

    # Executa o processo de automação
    automate_download(site_url, user_email, user_password)
