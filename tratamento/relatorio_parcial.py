import os
import time
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import logging
import subprocess

# Configurar logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class RelatorioParcial:

    def __init__(self, diretorio_download, usuario_inicial) -> None:
        pass

    def verificar_argumentos():
        if len(sys.argv) < 1:
            raise Exception("Erro: Login e senha de administrador não fornecidos.")
        usuario_inicial = sys.argv[1]
        print(usuario_inicial)
        
        return usuario_inicial

    def criar_diretorio_download():
        diretorio_download = os.path.join(os.getcwd(), "downloads parcial")
        if not os.path.exists(diretorio_download):
            os.makedirs(diretorio_download)
        return diretorio_download

    # Função para configurar o navegador
    def configurar_navegador(diretorio_download, headless=False):
        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": diretorio_download}
        chrome_options.add_experimental_option("prefs", prefs)
        
        if headless:
            chrome_options.add_argument("--headless")  # Executa o navegador em modo headless
            chrome_options.add_argument("--disable-gpu")  # Necessário para o modo headless funcionar corretamente em alguns sistemas
            chrome_options.add_argument("--window-size=1920,1080")  # Define um tamanho de janela padrão

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=chrome_options)
        return navegador

    def clicar_com_javascript(navegador, xpath, tempo_espera=0):
        try:
            elemento = WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            navegador.execute_script("arguments[0].click();", elemento)
            if tempo_espera > 0:
                time.sleep(tempo_espera)
        except Exception as e:
            logging.error(f"Erro ao clicar no elemento {xpath}: {e}")

    def fazer_login(self, navegador, usuario_inicial):
        max_retries = 3
        for attempt in range(max_retries):
            try:
                navegador.get("https://santacasabahia.mentorlearn.com/login")
                navegador.maximize_window()
                campo_user = WebDriverWait(navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="username"]'))
                )
                campo_user.send_keys(usuario_inicial)
                campo_senha = WebDriverWait(navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="password"]'))
                )
                campo_senha.send_keys(usuario_inicial)
                self.clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-login/div/div/div[2]/form/div/div[3]/p/button')
                WebDriverWait(navegador, 20).until(EC.url_changes("https://santacasabahia.mentorlearn.com/login"))
                return True
            except Exception as e:
                logging.error(f"Erro ao tentar fazer login com {usuario_inicial} na tentativa {attempt + 1}: {e}")
                time.sleep(5)  # Esperar alguns segundos antes de tentar novamente
        return False

    def gerar_relatorio_1(self, navegador, username, diretorio_download):
        try:
            navegador.get('https://santacasabahia.mentorlearn.com/reports/sim_typ8(secondary:reports-settings)')
            campo_data = WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-report-settings/div/mat-accordion/mat-expansion-panel[1]/div/div/div/div/mat-form-field[1]/div/div[1]/div[1]/input'))
            )
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.DELETE)
            campo_data.send_keys('1/1/2021')
            campo_data.send_keys(Keys.TAB)
            
            time.sleep(5)  # Espera fazer a busca de todos os itens

            if WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-row[1]'))):
                self.clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label', 2)
                # clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label', 2)
                self.clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-toolbar/div/mat-toolbar/div[5]/ml-toolbar-more-menu/button', 2)
                self.clicar_com_javascript(navegador, '/html/body/div/div[2]/div/div/div/button', 2)
                self.clicar_com_javascript(navegador, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/mat-checkbox/label', 2)
                self.clicar_com_javascript(navegador, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[2]/mat-radio-group/mat-radio-button[2]/label',2)
                self.clicar_com_javascript(navegador, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[3]/div[2]/button', 2)

                # Verificar até que o ícone de carregamento desapareça
                WebDriverWait(navegador, 240).until(
                    EC.invisibility_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[3]'))
                )
                time.sleep(10)  # Espera adicional após o desaparecimento do ícone

                # Verificação contínua da existência do arquivo baixado e renomeado
                tempo_limite = time.time() + 60  # 60 segundos de espera
                novo_nome = os.path.join(diretorio_download, f"{username}_report8.zip")
                while time.time() < tempo_limite:
                    try:
                        ultimo_arquivo = max([os.path.join(diretorio_download, f) for f in os.listdir(diretorio_download)], key=os.path.getctime)
                        os.rename(ultimo_arquivo, novo_nome)
                        break
                    except (ValueError, FileNotFoundError):
                        time.sleep(1)

                if not os.path.exists(novo_nome):
                    logging.error(f"Relatório 8 não encontrado para {username}. Verifique se o relatório foi gerado corretamente.")
            else:
                logging.warning(f"O elemento para selecionar todos os itens do relatório não foi encontrado para {username}. Não pode ser gerado o relatório 8.")

        except Exception as e:
            logging.error(f"Erro ao gerar o relatório 8 para {username}: {e}")

    # Função para gerar o segundo relatório
    def gerar_relatorio_2(self, navegador, usuario_inicial, diretorio_download):
        try:
            navegador.get('https://santacasabahia.mentorlearn.com/reports/sim_typ2(secondary:reports-settings)')
            campo_data = WebDriverWait(navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-report-settings/div/mat-accordion/mat-expansion-panel[1]/div/div/div/div/mat-form-field[1]/div/div[1]/div[1]/input'))
            )
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.DELETE)
            campo_data.send_keys('1/1/2021')
            campo_data.send_keys(Keys.TAB)
            
            time.sleep(5)  # Espera fazer a busca de todos os itens

            if WebDriverWait(navegador, 20).until(EC.presence_of_element_located((By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-row[1]'))):
                self.clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label', 2)
                self.clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label', 2)
                self.clicar_com_javascript(navegador, '/html/body/ml-app/div/ml-toolbar/div/mat-toolbar/div[5]/ml-toolbar-more-menu/button', 2)
                self.clicar_com_javascript(navegador, '/html/body/div/div[2]/div/div/div/button', 2)
                self.clicar_com_javascript(navegador, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/mat-checkbox/label', 2)
                self.clicar_com_javascript(navegador, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[2]/mat-radio-group/mat-radio-button[2]/label',2)
                self.clicar_com_javascript(navegador, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[3]/div[2]/button', 2)

                # Verificar até que o ícone de carregamento desapareça
                WebDriverWait(navegador, 240).until(
                    EC.invisibility_of_element_located((By.XPATH, '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[3]'))
                )
                time.sleep(10)  # Espera adicional após o desaparecimento do ícone

                # Verificação contínua da existência do arquivo baixado e renomeado
                tempo_limite = time.time() + 60  # 60 segundos de espera
                novo_nome = os.path.join(diretorio_download, f"{usuario_inicial}_report2.zip")
                while time.time() < tempo_limite:
                    try:
                        ultimo_arquivo = max([os.path.join(diretorio_download, f) for f in os.listdir(diretorio_download)], key=os.path.getctime)
                        os.rename(ultimo_arquivo, novo_nome)
                        break
                    except (ValueError, FileNotFoundError):
                        time.sleep(1)

                if not os.path.exists(novo_nome):
                    logging.error(f"Relatório 2 não encontrado para {usuario_inicial}. Verifique se o relatório foi gerado corretamente.")
            else:
                logging.warning(f"O elemento para selecionar todos os itens do relatório não foi encontrado para {usuario_inicial}. Não pode ser gerado o relatório 2.")

        except Exception as e:
            logging.error(f"Indo para o proximo relatorio")

    def processar_usuarios(self, navegador, usuario_inicial, diretorio_download):
        if usuario_inicial:
            logging.info(f"Iniciando geração de relatórios para {usuario_inicial}")
            self.gerar_relatorio_1(navegador, usuario_inicial, diretorio_download)
            self.gerar_relatorio_2(navegador, usuario_inicial, diretorio_download)
            logging.info(f"Relatórios gerados para {usuario_inicial}")

    def main(self):
        usuario_inicial = self.verificar_argumentos()
        diretorio_download = self.criar_diretorio_download()
        navegador = self.configurar_navegador(diretorio_download)

        if self.fazer_login(navegador, usuario_inicial):
            self.processar_usuarios(navegador, usuario_inicial, diretorio_download)
            navegador.quit()
            logging.info("Todos os relatórios foram gerados com sucesso.")
            
            # Inicia o TratamentoZIP.py
            logging.info("Iniciando o TratamentoZIP.py...")
            resultado_zip = subprocess.run(["python", "TratamentoZIPparcial.py"])

            if resultado_zip.returncode == 0:
                logging.info("TratamentoZIP.py concluído com sucesso.")
                
                # Agora inicia o TratamentoCSV.py
                logging.info("Iniciando o TratamentoCSVparcial.py...")
                resultado_csv = subprocess.run(["python", "TratamentoCSVparcial.py"])

                if resultado_csv.returncode == 0:
                    logging.info("TratamentoCSV.py concluído com sucesso.")
                else:
                    logging.error("TratamentoCSVparcial.py falhou.")
            else:
                logging.error("TratamentoZIP.py falhou. TratamentoCSVparcial.py não será iniciado.")

            logging.info("Processo completo.")
        else:
            logging.error("Falha ao fazer login. Verifique suas credenciais.")

    if __name__ == "__main__":
        main()