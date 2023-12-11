import pandas as pd
import openpyxl

# Carregue o arquivo Excel existente
workbook = openpyxl.load_workbook('planilhajobs.xlsx')

# Selecione a planilha ativa (a primeira por padrão)
sheet = workbook.active

# Abra o arquivo de nomes
with open("nomes_unicos.txt", "r") as file:
    nomes = file.read().splitlines()

# Loop através de cada nome no arquivo de nomes
for nome in nomes:
    # Carregue o arquivo CSV de setembro
    setembro_df = pd.read_excel("setembro.xlsx")

    # Carregue o arquivo CSV de outubro
    outubro_df = pd.read_excel("outubro.xlsx")

    # Crie uma lista de datas para pesquisar
    datas_setembro = ["10/09/23", "11/09/23", "12/09/23", "13/09/23", "14/09/23", "15/09/23", "16/09/23", "17/09/23"]
    datas_outubro = ["08/10/23", "09/10/23", "10/10/23", "11/10/23", "12/10/23", "13/10/23", "14/10/23", "15/10/23"]

    # Converta a coluna RUNSTARTTIMESTAMP para datetime
    setembro_df['RUNSTARTTIMESTAMP'] = pd.to_datetime(setembro_df['RUNSTARTTIMESTAMP'], format='%d/%m/%y %H:%M:%S')
    outubro_df['RUNSTARTTIMESTAMP'] = pd.to_datetime(outubro_df['RUNSTARTTIMESTAMP'], format='%d/%m/%y %H:%M:%S')

    # Consulte os dados de setembro e outubro e armazene em variáveis
    segundos_setembro = []
    segundos_outubro = []

    for data in datas_setembro:
        filtro_setembro = setembro_df[setembro_df['RUNSTARTTIMESTAMP'].dt.strftime('%d/%m/%y') == data]
        if not filtro_setembro.empty:
            segundos = filtro_setembro[setembro_df["JOBNAME"] == nome]["ELAPSEDRUNSECS"]
            segundos_setembro.extend(segundos.tolist())

    for data in datas_outubro:
        filtro_outubro = outubro_df[outubro_df['RUNSTARTTIMESTAMP'].dt.strftime('%d/%m/%y') == data]
        if not filtro_outubro.empty:
            segundos = filtro_outubro[outubro_df["JOBNAME"] == nome]["ELAPSEDRUNSECS"]
            segundos_outubro.extend(segundos.tolist())

    # Soma
    soma_setembro = sum(segundos_setembro)
    soma_outubro = sum(segundos_outubro)

    # Calculo outubro
    outsoma = soma_outubro
    outmedia = outsoma / 8
    outmediaarre = round(outmedia, 1)  # Média Arredondada
    outmedias = outmedia / 60
    outmediasemana = round(outmedias, 1)  # Média Semanal

    # Calculo setembro
    setsoma = soma_setembro
    setmedia = setsoma / 8
    setmediaarre = round(setmedia, 1)  # Média Arredondada
    setmedias = setmedia / 60
    setmediasemana = round(setmedias, 1)  # Média Semanal
    diferenca = outmediasemana - setmediasemana

    # Adicione suas variáveis na próxima linha vazia
    next_row = sheet.max_row + 1
    sheet[f'A{next_row}'] = nome
    sheet[f'B{next_row}'] = setmediasemana
    sheet[f'C{next_row}'] = outmediasemana
    sheet[f'D{next_row}'] = diferenca

# Salve o arquivo Excel após o loop
workbook.save('planilhajobs.xlsx')