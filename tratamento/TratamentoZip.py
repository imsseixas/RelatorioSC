import os
import zipfile
import pandas as pd

# Função para descompactar arquivos zip e pegar todos os arquivos da pasta 'all'
def descompactar_arquivos(diretorio_download):
    for filename in os.listdir(diretorio_download):
        if filename.endswith(".zip"):
            file_path = os.path.join(diretorio_download, filename)
            extract_dir = os.path.join(diretorio_download, filename.replace(".zip", ""))
            if not os.path.exists(extract_dir):
                os.makedirs(extract_dir)
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
                all_dir = os.path.join(extract_dir, 'all')
                if os.path.exists(all_dir):
                    for item in os.listdir(all_dir):
                        item_path = os.path.join(all_dir, item)
                        if os.path.isfile(item_path):
                            os.rename(item_path, os.path.join(extract_dir, item))

# Função para combinar arquivos CSV
def combinar_csvs(diretorio_download, output_file):
    data_frames = []
    for root, dirs, files in os.walk(diretorio_download):
        for file in files:
            if file.endswith(".csv"):
                file_path = os.path.join(root, file)
                df = pd.read_csv(file_path)
                data_frames.append(df)

    if data_frames:
        relatorio_geral = pd.concat(data_frames, ignore_index=True)
        relatorio_geral.to_csv(output_file, index=False)
        print(f"Arquivo combinado salvo como {output_file}")
    else:
        print("Nenhum arquivo CSV encontrado para combinar.")

# Função principal
def main():
    diretorio_download = os.path.join(os.getcwd(), "downloads")

    # Descompactar arquivos zip
    descompactar_arquivos(diretorio_download)

    # Combinar arquivos CSV
    combinar_csvs(diretorio_download, "dados.csv")

if __name__ == "__main__":
    main()
