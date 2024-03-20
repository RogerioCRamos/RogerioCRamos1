import pandas as pd
import openpyxl as op
import os

pasta = "2024-02 0118018 - Lista de Saldo aberto por Fornecedor 212002.xlsx"
caminho = rf"C:\Users\rcramos2\Desktop\Automacao INV\Tratamento saldo aberto\planilha"
absoluto = os.path.join(caminho,pasta)

df = pd.read_excel(absoluto, sheet_name=0, header=None)

df[0].ffill(inplace=True)
df[1].ffill(inplace=True)
df[3].ffill(inplace=True)

df.reset_index(inplace=True, drop=True)


indice_conta = df[df[4] == "Fatura"].index[0]
novo_cabecalho = df.iloc[indice_conta]
df.columns = novo_cabecalho


lista_index = [11,13,14,15,16,17,18]

df = df.drop(lista_index, axis='index')

valores1 = df.iloc[0:3, 2].values

# Atribuindo esses valores aos mesmos índices na coluna 'A'
df.iloc[0:3, 0] = valores1

lista_index2 = [3,5,6]

df = df.drop(lista_index2, axis='index')

valores2 = df.iloc[3:6, 2].values
df.iloc[0:3, 8] = valores2

df.loc[0:6, 2] = None

df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)

lista_index3 = [7,8]

df = df.drop(lista_index3, axis='index')

nome_arquivo = "Teste"

df.to_excel(f"{caminho}\{nome_arquivo}.xlsx",index=False)

workbook = op.load_workbook(f"{caminho}\{nome_arquivo}.xlsx")

sheet = workbook.active




...