import subprocess
import threading
import psutil
from tratamento.relatorio_parcial import RelatorioParcial

class ProcessManager:
    def __init__(self):
        self.process = None

    def run_script(self, script_name, *args):
        """Run a script in a separate thread."""
        def execute():
            try:
                self.process = subprocess.Popen(['python', script_name, *args])
                self.process.wait()
            except subprocess.CalledProcessError as e:
                print(f"Error executing {script_name}: {e}")

        thread = threading.Thread(target=execute)
        thread.start()

    def stop_process(self):
        """Stop the currently running process."""
        if self.process:
            try:
                parent = psutil.Process(self.process.pid)
                for child in parent.children(recursive=True):
                    child.terminate()
                parent.terminate()
                self.process = None
            except psutil.NoSuchProcess:
                print("No process found to stop.")
            except Exception as e:
                print(f"Error stopping the process: {e}")

    def run_script(self, script, *args):
        if script == './tratamento/relatorioSC.py':
            relatorio_sc = RelatorioSC("downloads", args[0])
            if relatorio_sc.fazer_login():
                relatorio_sc.gerar_relatorio_1()
                relatorio_sc.gerar_relatorio_2()
        elif script == './tratamento/relatorio_parcial.py':
            relatorio_parcial = RelatorioParcial("downloads par", args[0])
            if relatorio_parcial.fazer_login():
                relatorio_parcial.gerar_relatorio_parcial()
        elif script == './tratamento/tratamentoZIP.py':
            zip_processor = TratamentoZIP()
            zip_processor.executar()
        elif script == './tratamento/tratamentoCSV.py':
            csv_processor = TratamentoCSV()
            csv_processor.executar()

    def stop_process(self):
        for processo in self.processos:
            processo.terminate()
        self.processos = []