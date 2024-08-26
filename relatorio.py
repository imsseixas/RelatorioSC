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
            diretorio_download = "Downloads"
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