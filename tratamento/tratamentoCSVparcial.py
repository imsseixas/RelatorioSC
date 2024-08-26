import pandas as pd
from datetime import datetime, timedelta
import sys

# Verificar se o argumento do usuário foi passado
if len(sys.argv) < 2:
    print("Erro: o nome do usuário inicial deve ser passado como argumento.")
    sys.exit(1)

# Capturar o usuário inicial a partir dos argumentos da linha de comando
usuario_inicial = sys.argv[1]

# Carregar o arquivo CSV
df = pd.read_csv('dados.csv', low_memory=False)

# Remover colunas indesejadas
df = df.drop(columns=['Case Number', 'Unique Session Id', 'Repetition'])
df = df.iloc[:, :6]  # Manter apenas as primeiras 6 colunas

# Remover linhas duplicadas com base em todas as colunas
df = df.drop_duplicates()

# Imprimir o número de linhas após remover duplicatas
print(f"Número de linhas após remover duplicatas: {len(df)}")

# Salvar o DataFrame limpo em um novo arquivo CSV
df.to_csv('dados_limpos.csv', index=False)

# Concatenar colunas 'First Name' e 'Last Name' em 'Nome Completo'
df['Nome Completo'] = df['First Name'] + ' ' + df['Last Name']

# Atualizar a lista de matérias para incluir a nova matéria
materias = ['Robotic Basic Skills', 'FRS', 'Robotic Essential Skills', 'Robotic Suturing', 'Stapler']

# Adicionar coluna 'Surgical Modules'
df['Surgical Modules'] = df.apply(
    lambda row: 'Surgical Modules' if pd.isna(row['Case Name']) or not any(materia.lower() in row['Case Name'].lower() for materia in materias) else '',
    axis=1
)

# Função para converter hora no formato especificado
def converter_hora(hora_str):
    try:
        hora = hora_str.strip().split(' ')[-1]  # Pega a parte do tempo (por exemplo, "3:42PM")
        hora = hora[:-2]  # Remove AM/PM
        hora_24 = datetime.strptime(hora, '%I:%M').time()
        return hora_24
    except Exception as e:
        print(f"Erro ao converter hora: {hora_str}, erro: {e}")
        return None

# Função para calcular a diferença de tempo em segundos
def calcular_diferenca_tempo(row):
    try:
        start_time_str = row['Start Time']
        end_time_str = row['End Time']

        # Converter para formato de hora
        start_time = converter_hora(start_time_str)
        end_time = converter_hora(end_time_str)

        # Verificar se start_time e end_time são válidos
        if start_time is None or end_time is None:
            return 0

        # Converter para objetos datetime para calcular a diferença
        start_time = datetime.combine(datetime.min, start_time)
        end_time = datetime.combine(datetime.min, end_time)

        # Calcular a diferença de tempo
        diff = end_time - start_time
        if diff < timedelta(0):  # Se a diferença for negativa, não adicionar um dia
            return 0
        return diff.total_seconds()
    except Exception as e:
        print(f"Erro ao processar linha: {row}, erro: {e}")
        return 0

# Função para converter segundos para horas formatadas
def converter_segundos_para_horas(segundos):
    horas = int(segundos // 3600)
    minutos = int((segundos % 3600) // 60)
    return f"{horas}h {minutos}m"

# Função para calcular o tempo gasto em cada matéria por aluno linha por linha
def calcular_tempo_por_aluno(aluno):
    aluno_df = df[df['Nome Completo'] == aluno]
    soma_materias = {materia: 0 for materia in materias}
    soma_materias['Surgical Modules'] = 0
    soma_total = 0

    for _, row in aluno_df.iterrows():
        diff = calcular_diferenca_tempo(row)
        soma_total += diff  # Adiciona o tempo total
        materia_encontrada = False
        for materia in materias:
            if materia.lower() in str(row['Case Name']).lower():
                soma_materias[materia] += diff
                materia_encontrada = True
                break
        if not materia_encontrada:
            soma_materias['Surgical Modules'] += diff

    # Formatar os resultados em horas
    resultados_formatados = {materia: converter_segundos_para_horas(soma_materias[materia]) for materia in soma_materias}
    resultados_formatados['Nome Completo'] = aluno
    resultados_formatados['Tempo Total'] = converter_segundos_para_horas(soma_total)

    return resultados_formatados

# Listar todos os alunos únicos
alunos_unicos = df['Nome Completo'].unique()

# Gerar os resultados para todos os alunos
resultados = [calcular_tempo_por_aluno(aluno) for aluno in alunos_unicos]

# Converter para DataFrame para exportar para XLSX
df_resultados = pd.DataFrame(resultados)

# Salvar o DataFrame em um arquivo XLSX com o nome do usuário inicial
nome_arquivo = f'Resultados_{usuario_inicial}.xlsx'
df_resultados.to_excel(nome_arquivo, index=False)

print(f"Processamento concluído. Resultados salvos em '{nome_arquivo}'.")
