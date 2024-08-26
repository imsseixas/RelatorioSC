import os
import tkinter as tk
from tkinter import ttk, messagebox
from process_manager import ProcessManager
from file_manager import save_credentials, load_credentials, load_users
from tkinter import Entry
import os
import time
import sys
import openpyxl
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import logging
import subprocess

class RelatorioBase:
    def __init__(self, diretorio_download, usuario_inicial):
        self.diretorio_download = diretorio_download
        self.usuario_inicial = usuario_inicial
        self.navegador = self.configurar_navegador()

    def configurar_navegador(self, headless=False):
        chrome_options = webdriver.ChromeOptions()
        prefs = {"download.default_directory": self.diretorio_download}
        chrome_options.add_experimental_option("prefs", prefs)
        
        if headless:
            chrome_options.add_argument("--headless")
            chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--window-size=1920,1080")

        servico = Service(ChromeDriverManager().install())
        navegador = webdriver.Chrome(service=servico, options=chrome_options)
        return navegador

    def clicar_com_javascript(self, xpath, tempo_espera=0):
        try:
            elemento = WebDriverWait(self.navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            self.navegador.execute_script("arguments[0].click();", elemento)
            if tempo_espera > 0:
                time.sleep(tempo_espera)
        except Exception as e:
            logging.error(f"Erro ao clicar no elemento {xpath}: {e}")

    def fazer_login(self):
        max_retries = 3
        for attempt in range(max_retries):
            try:
                self.navegador.get("https://santacasabahia.mentorlearn.com/login")
                self.navegador.maximize_window()
                campo_user = WebDriverWait(self.navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="username"]'))
                )
                campo_user.send_keys(self.usuario_inicial)
                campo_senha = WebDriverWait(self.navegador, 20).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="password"]'))
                )
                campo_senha.send_keys(self.usuario_inicial)
                self.clicar_com_javascript('/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-login/div/div/div[2]/form/div/div[3]/p/button')
                WebDriverWait(self.navegador, 20).until(EC.url_changes("https://santacasabahia.mentorlearn.com/login"))
                return True
            except Exception as e:
                logging.error(f"Erro ao tentar fazer login com {self.usuario_inicial} na tentativa {attempt + 1}: {e}")
                time.sleep(5)
        return False

    def gerar_relatorio(self, url, xpath_elementos, nome_arquivo):
        try:
            self.navegador.get(url)
            campo_data = WebDriverWait(self.navegador, 20).until(
                EC.presence_of_element_located((By.XPATH, xpath_elementos['campo_data']))
            )
            campo_data.send_keys(Keys.CONTROL + "a")
            campo_data.send_keys(Keys.DELETE)
            campo_data.send_keys('1/1/2021')
            campo_data.send_keys(Keys.TAB)
            
            time.sleep(5)

            if WebDriverWait(self.navegador, 20).until(EC.presence_of_element_located((By.XPATH, xpath_elementos['elemento_lista']))):
                for xpath in xpath_elementos['xpaths']:
                    self.clicar_com_javascript(xpath, 2)
                WebDriverWait(self.navegador, 240).until(
                    EC.invisibility_of_element_located((By.XPATH, xpath_elementos['elemento_carregamento']))
                )
                time.sleep(10)

                tempo_limite = time.time() + 60
                novo_nome = os.path.join(self.diretorio_download, nome_arquivo)
                while time.time() < tempo_limite:
                    try:
                        ultimo_arquivo = max([os.path.join(self.diretorio_download, f) for f in os.listdir(self.diretorio_download)], key=os.path.getctime)
                        os.rename(ultimo_arquivo, novo_nome)
                        break
                    except (ValueError, FileNotFoundError):
                        time.sleep(1)

                if not os.path.exists(novo_nome):
                    logging.error(f"{nome_arquivo} não encontrado para {self.usuario_inicial}. Verifique se o relatório foi gerado corretamente.")
            else:
                logging.warning(f"O elemento para selecionar todos os itens do relatório não foi encontrado para {self.usuario_inicial}. Não pode ser gerado o relatório.")
        except Exception as e:
            logging.error(f"Erro ao gerar o relatório para {self.usuario_inicial}: {e}")

class RelatorioSC(RelatorioBase):
    def gerar_relatorio_1(self):
        url = 'https://santacasabahia.mentorlearn.com/reports/sim_typ1(secondary:reports-settings)'
        xpaths = [
            '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label',
            '/html/body/ml-app/div/ml-toolbar/div/mat-toolbar/div[5]/ml-toolbar-more-menu/button',
            '/html/body/div/div[2]/div/div/div/button',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/mat-checkbox/label',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[2]/mat-radio-group/mat-radio-button[2]/label',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[3]/div[2]/button'
        ]
        xpath_elementos = {
            'campo_data': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-report-settings/div/mat-accordion/mat-expansion-panel[1]/div/div/div/div/mat-form-field[1]/div/div[1]/div[1]/input',
            'elemento_lista': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-row[1]',
            'xpaths': xpaths,
            'elemento_carregamento': '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[3]'
        }
        self.gerar_relatorio(url, xpath_elementos, f"{self.usuario_inicial}_report1.zip")

    def gerar_relatorio_2(self):
        url = 'https://santacasabahia.mentorlearn.com/reports/sim_typ2(secondary:reports-settings)'
        xpaths = [
            '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label',
            '/html/body/ml-app/div/ml-toolbar/div/mat-toolbar/div[5]/ml-toolbar-more-menu/button',
            '/html/body/div/div[2]/div/div/div/button',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/mat-checkbox/label',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[2]/mat-radio-group/mat-radio-button[2]/label',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[3]/div[2]/button'
        ]
        xpath_elementos = {
            'campo_data': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-report-settings/div/mat-accordion/mat-expansion-panel[1]/div/div/div/div/mat-form-field[1]/div/div[1]/div[1]/input',
            'elemento_lista': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-row[1]',
            'xpaths': xpaths,
            'elemento_carregamento': '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[3]'
        }
        self.gerar_relatorio(url, xpath_elementos, f"{self.usuario_inicial}_report2.zip")

class RelatorioParcial(RelatorioBase):
    def gerar_relatorio_parcial(self):
        url = 'https://santacasabahia.mentorlearn.com/reports/sim_typ1(secondary:reports-settings)'
        xpaths = [
            '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label',
            '/html/body/ml-app/div/ml-toolbar/div/mat-toolbar/div[5]/ml-toolbar-more-menu/button',
            '/html/body/div/div[2]/div/div/div/button',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/mat-checkbox/label',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[2]/mat-radio-group/mat-radio-button[2]/label',
            '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[3]/div[2]/button'
        ]
        xpath_elementos = {
            'campo_data': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-report-settings/div/mat-accordion/mat-expansion-panel[1]/div/div/div/div/mat-form-field[1]/div/div[1]/div[1]/input',
            'elemento_lista': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-row[1]',
            'xpaths': xpaths,
            'elemento_carregamento': '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[3]'
        }
        self.gerar_relatorio(url, xpath_elementos, f"{self.usuario_inicial}_report_partial.zip")

        def gerar_relatorio_2(self):
            url = 'https://santacasabahia.mentorlearn.com/reports/sim_typ2(secondary:reports-settings)'
            xpaths = [
                '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-header-row/mat-header-cell[1]/mat-checkbox/label',
                '/html/body/ml-app/div/ml-toolbar/div/mat-toolbar/div[5]/ml-toolbar-more-menu/button',
                '/html/body/div/div[2]/div/div/div/button',
                '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/mat-checkbox/label',
                '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[2]/mat-radio-group/mat-radio-button[2]/label',
                '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[3]/div[2]/button'
            ]
            xpath_elementos = {
                'campo_data': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav[2]/div/ng-component/ml-report-settings/div/mat-accordion/mat-expansion-panel[1]/div/div/div/div/mat-form-field[1]/div/div[1]/div[1]/input',
                'elemento_lista': '/html/body/ml-app/div/ml-main-content/mat-sidenav-container/mat-sidenav-content/ml-reports/div/mat-table/mat-row[1]',
                'xpaths': xpaths,
                'elemento_carregamento': '/html/body/div[1]/div[2]/div/mat-dialog-container/ml-dynamic-dialog/div/div[2]/ml-export-performance/div/div[3]'
            }
            self.gerar_relatorio(url, xpath_elementos, f"{self.usuario_inicial}_report2.zip")

        def main():
            diretorio_download = "c:\\Downloads"
            usuario_inicial = "usuario_teste"

            relatorio_sc = RelatorioSC(diretorio_download, usuario_inicial)
            if relatorio_sc.fazer_login():
                relatorio_sc.gerar_relatorio_1()
                relatorio_sc.gerar_relatorio_2()

            relatorio_parcial = RelatorioParcial(diretorio_download, usuario_inicial)
            if relatorio_parcial.fazer_login():
                relatorio_parcial.gerar_relatorio_parcial()

class TratamentoZIP:
    def __init__(self):
        pass

    def executar(self):
        # Implementar a lógica do script TratamentoZIP.py aqui
        print("Executando TratamentoZIP")
        # Coloque aqui o código necessário para o tratamento de ZIP

class TratamentoCSV:
    def __init__(self):
        pass

    def executar(self):
        # Implementar a lógica do script TratamentoCSV.py aqui
        print("Executando TratamentoCSV")
        # Coloque aqui o código necessário para o tratamento de CSV

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Tela de Acesso aos Relatórios")
        self.root.geometry("600x400")
        self.root.configure(bg="#575e62")
        self.centralizar_janela()

        self.process_manager = ProcessManager()

        self.create_widgets()

    def centralizar_janela(self):
        self.root.update_idletasks()
        largura = self.root.winfo_width()
        altura = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (largura // 2)
        y = (self.root.winfo_screenheight() // 2) - (altura // 2)
        self.root.geometry(f'{largura}x{altura}+{x}+{y}')

    def create_widgets(self):
        # Configurações de fonte e estilo
        font_style = ("Arial", 12)

        # Estilo dos botões
        style = ttk.Style()
        style.configure("TButton", font=font_style, padding=10)

        # Criar o notebook (abas)
        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=True, fill='both')

        # Aba Geral
        aba_geral = ttk.Frame(notebook)
        notebook.add(aba_geral, text='Relatório Geral')

        # Aba Relatório Parcial
        aba_relatorio = ttk.Frame(notebook)
        notebook.add(aba_relatorio, text='Relatório Parcial')

        # Frame para a seção de login na aba Geral
        frame_login = tk.Frame(aba_geral, bg="#ffffff", padx=20, pady=20, relief=tk.RAISED, bd=2)
        frame_login.pack(side="left", expand=True, fill="both")

        # Frame para a seção do botão na aba Geral
        frame_button = tk.Frame(aba_geral, bg="#ffffff", padx=20, pady=20, relief=tk.RAISED, bd=2)
        frame_button.pack(side="right", expand=True, fill="both")

        # Texto acima do login na aba Geral
        tk.Label(frame_login, text="Entrar como administrador", bg="#ffffff", font=("Arial", 14, "bold")).pack(pady=10)

        # Campos para login e senha no frame da esquerda
        tk.Label(frame_login, text="Login:", bg="#ffffff", font=font_style).pack(pady=5, anchor="w")
        self.entry_username = ttk.Entry(frame_login, font=font_style, width=30)
        self.entry_username.pack(pady=5)

        tk.Label(frame_login, text="Senha:", bg="#ffffff", font=font_style).pack(pady=5, anchor="w")
        self.entry_password = ttk.Entry(frame_login, show="*", font=font_style, width=30)
        self.entry_password.pack(pady=5)

        # Carregar credenciais salvas (se existirem)
        login_salvo, senha_salva = load_credentials()
        self.entry_username.insert(0, login_salvo)
        self.entry_password.insert(0, senha_salva)

        # Botões no frame da direita na aba Geral
        btn_run_relatorioSC = ttk.Button(frame_button, text="Executar relatorio", command=self.run_relatorioSC)
        btn_run_relatorioSC.pack(pady=10)

        btn_stop = ttk.Button(frame_button, text="Parar Ações", command=self.stop_actions)
        btn_stop.pack(side="bottom", pady=10)

        # Botão para abrir a planilha na aba Geral
        self.btn_open_planilha_geral = ttk.Button(frame_button, text="Abrir Planilha Geral", command=self.abrir_planilha_geral)
        self.btn_open_planilha_geral.pack(pady=10)
        self.atualizar_estado_botao_geral()  # Atualizar estado inicial

        # Aba de Relatório Parcial
        tk.Label(aba_relatorio, text="Selecione um usuário para o relatório", font=("Arial", 14, "bold")).pack(pady=20)

        self.entry_usuario = Entry(aba_relatorio)
        self.entry_usuario.pack()

        try:
            usuarios = load_users()  # Tenta carregar os usuários da planilha
            if usuarios:  # Se os usuários forem carregados com sucesso
                # Remove o campo de entrada anterior
                self.entry_usuario.pack_forget()

                # Cria um combobox com os usuários carregados
                self.combo_usuario = ttk.Combobox(aba_relatorio, values=usuarios, state="normal")
                self.combo_usuario.pack()
        except Exception as e:
            print(f"Erro ao carregar usuários: {e}")

        btn_run_relatorio_parcial = ttk.Button(aba_relatorio, text="Executar Relatório Parcial", command=self.run_relatorio_parcial)
        btn_run_relatorio_parcial.pack(pady=10)

        # Botão para abrir a planilha na aba Relatório Parcial
        self.btn_open_planilha_parcial = ttk.Button(aba_relatorio, text="Abrir Planilha Parcial", command=self.abrir_planilha_parcial)
        self.btn_open_planilha_parcial.pack(pady=10)
        self.btn_open_planilha_parcial.config(state="disable")  # Desabilitar inicialmente

    def run_relatorioSC(self):
        username = self.entry_username.get()
        password = self.entry_password.get()

        if not username or not password:
            messagebox.showwarning("Atenção", "Por favor, insira o login e senha de administrador.")
            return

        save_credentials(username, password)

        relatorio = RelatorioSC(username, password)
        relatorio.executar()

        # Habilitar o botão para abrir a planilha na aba Geral
        self.atualizar_estado_botao_geral()

    def stop_actions(self):
        self.process_manager.stop_process()
        messagebox.showinfo("Ação", "Ações paradas com sucesso.")

    def run_relatorio_parcial(self):
        usuario_selecionado = self.combo_usuario.get()
        if not usuario_selecionado:
            messagebox.showwarning("Atenção", "Por favor, selecione um usuário.")
            return

        relatorio_parcial = RelatorioParcial(usuario_selecionado)
        relatorio_parcial.executar()

        # Habilitar o botão para abrir a planilha na aba Relatório Parcial
        self.btn_open_planilha_parcial.config(state="normal")

    def abrir_planilha_geral(self):
        # Abrir a planilha do relatório geral
        if os.path.exists("Resultados.xlsx"):
            os.startfile("Resultados.xlsx")
        else:
            messagebox.showwarning("Atenção", "A planilha geral não existe.")

    def abrir_planilha_parcial(self):
        usuario_selecionado = self.combo_usuario.get()
        if not usuario_selecionado:
            messagebox.showwarning("Atenção", "Nenhum usuário selecionado.")
            return

        # Abrir a planilha do relatório parcial
        arquivo = f'Resultados_{usuario_selecionado}.xlsx'
        if os.path.exists(arquivo):
            os.startfile(arquivo)
        else:
            messagebox.showwarning("Atenção", f"A planilha para o usuário '{usuario_selecionado}' não existe.")

    def atualizar_estado_botao_geral(self):
        if os.path.exists("Resultados.xlsx"):
            self.btn_open_planilha_geral.config(state="normal")
        else:
            self.btn_open_planilha_geral.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
