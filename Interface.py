import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import threading
import os
import psutil
import pandas as pd

# Função para centralizar a janela na tela
def centralizar_janela(janela):
    janela.update_idletasks()
    largura = janela.winfo_width()
    altura = janela.winfo_height()
    x = (janela.winfo_screenwidth() // 2) - (largura // 2)
    y = (janela.winfo_screenheight() // 2) - (altura // 2)
    janela.geometry(f'{largura}x{altura}+{x}+{y}')

# Função para ler o login e senha do arquivo
def carregar_credenciais():
    if os.path.exists("credenciais.txt"):
        with open("credenciais.txt", "r") as file:
            linhas = file.readlines()
            if len(linhas) >= 2:
                return linhas[0].strip(), linhas[1].strip()
    return "", ""

# Função para salvar o login e senha no arquivo
def salvar_credenciais(username, password):
    with open("credenciais.txt", "w") as file:
        file.write(f"{username}\n{password}")

# Variável global para armazenar o processo
processo = None

# Função para executar o relatorioSC.py
def run_relatorioSC():
    global processo
    username = entry_username.get()
    password = entry_password.get()
    
    if not username or not password:
        messagebox.showwarning("Atenção", "Por favor, insira o login e senha de administrador.")
        return

    salvar_credenciais(username, password)  # Salvar as credenciais

    def execute_script():
        global processo
        try:
            processo = subprocess.Popen(['python', 'relatorioSC.py', username, password])
            processo.wait()  # Espera o processo terminar
            messagebox.showinfo("Sucesso", "relatorioSC executado com sucesso!")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Erro", f"Falha ao executar relatorioSC: {e}")

    thread = threading.Thread(target=execute_script)
    thread.start()

# Função para parar as ações
def stop_actions():
    global processo
    if processo:
        try:
            process_id = processo.pid
            parent = psutil.Process(process_id)
            for child in parent.children(recursive=True):
                child.terminate()
            parent.terminate()
            processo = None
            messagebox.showinfo("Ação", "Ações paradas com sucesso.")
        except psutil.NoSuchProcess:
            messagebox.showwarning("Ação", "Nenhum processo encontrado para parar.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao tentar parar o processo: {e}")
    else:
        messagebox.showwarning("Ação", "Nenhuma ação está em andamento.")

# Função para carregar os usuários do arquivo Excel
def carregar_usuarios():
    if os.path.exists("usuarios.xlsx"):
        try:
            df = pd.read_excel("usuarios.xlsx")
            if "Usuario" in df.columns:
                usuarios = df["Usuario"].tolist()
                return usuarios
            else:
                messagebox.showwarning("Atenção", "Coluna 'Usuario' não encontrada no arquivo Excel.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao ler o arquivo Excel: {e}")
    else:
        messagebox.showwarning("Atenção", "Arquivo 'usuarios.xlsx' não encontrado.")
    return []

# Função para executar o relatorioSC.py com um usuário específico
def run_relatorio_parcial():
    usuario_selecionado = combo_usuario.get()
    if not usuario_selecionado:
        messagebox.showwarning("Atenção", "Por favor, selecione um usuário.")
        return

    def execute_script():
        try:
            subprocess.run(['python', 'relatorioSC.py', usuario_selecionado], check=True)
            messagebox.showinfo("Sucesso", "Relatório parcial executado com sucesso!")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Erro", f"Falha ao executar relatorioSC: {e}")

    thread = threading.Thread(target=execute_script)
    thread.start()

# Configuração da janela principal
root = tk.Tk()
root.title("Tela de Acesso a Programas")
root.geometry("600x400")
root.configure(bg="#575e62")
centralizar_janela(root)

# Estilo de fonte
font_style = ("Arial", 12)

# Estilo dos botões
button_style = {
    'font': font_style,
    'bg': "#4CAF50",
    'fg': "white",
    'padx': 10,
    'pady': 5,
    'relief': tk.RAISED,
    'bd': 2
}

# Função para estilizar botões com arredondamento
def estilizar_botao(botao, cor="#4CAF50"):
    botao.config(bg=cor, fg="white", padx=10, pady=5, relief=tk.RAISED, bd=2)
    botao.bind("<Enter>", lambda e: botao.config(bg="#45a049"))
    botao.bind("<Leave>", lambda e: botao.config(bg=cor))

# Criar o notebook (abas)
notebook = ttk.Notebook(root)
notebook.pack(expand=True, fill='both')

# Aba Geral
aba_geral = ttk.Frame(notebook)
notebook.add(aba_geral, text='Geral')

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
entry_username = tk.Entry(frame_login, font=font_style, width=30)
entry_username.pack(pady=5)

tk.Label(frame_login, text="Senha:", bg="#ffffff", font=font_style).pack(pady=5, anchor="w")
entry_password = tk.Entry(frame_login, show="*", font=font_style, width=30)
entry_password.pack(pady=5)

# Carregar credenciais salvas (se existirem)
login_salvo, senha_salva = carregar_credenciais()
entry_username.insert(0, login_salvo)
entry_password.insert(0, senha_salva)

# Botões no frame da direita na aba Geral
btn_run_relatorioSC = tk.Button(frame_button, text="Executar relatorioSC", command=run_relatorioSC)
estilizar_botao(btn_run_relatorioSC)
btn_run_relatorioSC.pack(pady=10)

btn_stop = tk.Button(frame_button, text="Parar Ações", command=stop_actions)
estilizar_botao(btn_stop, cor="#f44336")
btn_stop.pack(side="bottom", pady=10)

# Aba de Relatório Parcial
usuarios = carregar_usuarios()
tk.Label(aba_relatorio, text="Selecione um usuário para o relatório", font=("Arial", 14, "bold")).pack(pady=20)

combo_usuario = ttk.Combobox(aba_relatorio, values=usuarios, state="readonly")
combo_usuario.pack(pady=10)

btn_run_relatorio_parcial = tk.Button(aba_relatorio, text="Executar Relatório Parcial", command=run_relatorio_parcial)
estilizar_botao(btn_run_relatorio_parcial)
btn_run_relatorio_parcial.pack(pady=10)

# Iniciar a aplicação
root.mainloop()
