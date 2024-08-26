import os
import pandas as pd

def save_credentials(username, password):
    with open("credenciais.txt", "w") as file:
        file.write(username + "\n" + password)

def load_credentials():
    if os.path.exists("credenciais.txt"):
        with open("credenciais.txt", "r") as file:
            lines = file.readlines()
            if len(lines) >= 2:
                username = lines[0].strip()
                password = lines[1].strip()
                return username, password
    return "", ""

# Função para carregar usuários do Excel
def load_users():
    if os.path.exists("usuarios.xlsx"):
        try:
            df = pd.read_excel("usuarios.xlsx")
            if "loginUser" in df.columns:
                return df["loginUser"].tolist()
            else:
                raise ValueError("Coluna 'Usuario' não encontrada.")
        except Exception as e:
            raise e
    else:
        raise FileNotFoundError("Arquivo 'usuarios.xlsx' não encontrado.")
    
def open_excel():
    if os.path.exists("Resultados.xlsx"):
        os.startfile("Resultados.xlsx")
    else:
        raise FileNotFoundError("Arquivo Resultados não encontrado")
