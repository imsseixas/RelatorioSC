from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import time
import openpyxl

def clicar_com_javascript(navegador, xpath, tempo_espera=0):
    try:
        elemento = WebDriverWait(navegador, 20).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        navegador.execute_script("arguments[0].click();", elemento)
        if tempo_espera > 0:
            time.sleep(tempo_espera)
    except Exception as e:
        print(f"Erro ao clicar no elemento {xpath}")

# Instalar o driver do Chrome
servico = Service(ChromeDriverManager().install())

# Inicializar o navegador
navegador = webdriver.Chrome(service=servico)

# Navegar para a página do Santa Casa
URL = "https://santacasabahia.mentorlearn.com/login"
navegador.get(URL)
navegador.maximize_window()

# Esperar e preencher os campos de usuário e senha
campo_user = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="username"]'))
)
campo_user.send_keys('RobsonS')

campo_senha = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, '//*[@id="password"]'))
)
campo_senha.send_keys('RobsonS')
campo_senha.send_keys(Keys.RETURN)

# Adicionar um tempo de espera para garantir que a página de login tenha carregado
time.sleep(5)

# Abre o link dos grupos
navegador.get('https://santacasabahia.mentorlearn.com/users-groups(secondary:users-groups-settings)')

# Expande a pasta do PRONON
expandir_pronon = WebDriverWait(navegador, 20).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-users-groups/div/div/ml-users-groups-view/div/mat-tree/mat-tree-node[4]/div/button'))
)
expandir_pronon.click()

# Adicionar um tempo de espera para garantir que a pasta PRONON foi expandida
time.sleep(10)

# Criar um novo arquivo Excel e adicionar um cabeçalho
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Dados"
sheet.append(["Nome", "Progresso"])

# Função para extrair informações necessárias e adicionar ao Excel
def extrair_informacoes():
    try:
        nome = navegador.find_element(By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-users-groups-settings/div/div').text
    except:
        nome = "N/A"
    
    try:
        progresso = navegador.find_element(By.XPATH, '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-users-groups-settings/mat-tab-group/div/mat-tab-body[2]/div/ml-users-groups-summary/div/ml-course-list/ml-course-item/mat-expansion-panel/mat-expansion-panel-header/span[1]/div/div/div[2]').text
    except:
        progresso = "N/A"

    # Adicionar os dados na planilha
    sheet.append([nome, progresso])

    # Salvar o arquivo Excel após cada extração de informações
    workbook.save("dados.xlsx")

# Loop para iterar sobre as diferentes pastas
for i in range(5, 220):  # Ajuste o range conforme necessário
    # Expandir a pasta atual
    expandir_xpath = f'/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-users-groups/div/div/ml-users-groups-view/div/mat-tree/mat-tree-node[{i}]'
    
    try:
        clicar_com_javascript(navegador, expandir_xpath)
        time.sleep(2)  # Espera para garantir que o conteúdo seja carregado

        # Clicar no sumário
        sumario_xpath = f'/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-users-groups-settings/mat-tab-group/mat-tab-header/div[2]/div/div/div[2]'
        sumario = WebDriverWait(navegador, 20).until(
            EC.presence_of_element_located((By.XPATH, sumario_xpath))
        )
        sumario.click()
        time.sleep(2)  # Espera para garantir que o conteúdo seja carregado

        # Chamar a função para extrair informações
        extrair_informacoes()

        # Desmarcar a pasta atual
        clicar_com_javascript(navegador, expandir_xpath)
        time.sleep(1)

    except Exception as e:
        print(f"Erro ao expandir ou manipular a pasta {i}: {e}")
        continue

# Fechar o navegador
navegador.quit()
