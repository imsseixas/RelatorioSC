import os
import tkinter as tk
from tkinter import ttk, messagebox
from process_manager import ProcessManager
from file_manager import save_credentials, load_credentials, load_users
from tratamento import *
from tkinter import Entry



class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Tela de Acesso a os Relatórios")
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

        self.process_manager.run_script('./tratamento/relatorioSC.py', username, password)

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

        self.process_manager.run_script('./tratamento/relatorio_parcial.py', usuario_selecionado)

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
