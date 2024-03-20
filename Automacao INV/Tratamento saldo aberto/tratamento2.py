import pandas as pd
import openpyxl as op
import os

arquivo = "2024-02 0118018 - Lista de Saldo aberto por Fornecedor 212002.xlsx"
caminho = rf"C:\Users\rcramos2\Desktop\Automacao INV\Tratamento saldo aberto\planilha"
absoluto = os.path.join(caminho,arquivo)

df = pd.read_excel(absoluto, sheet_name=0, header=None)

df[0].ffill(inplace=True)
df[1].ffill(inplace=True)
df[3].ffill(inplace=True)

df.reset_index(inplace=True, drop=True)

cabeçalho = df.iloc[0:9, 2]


indice_conta = df[df[0] == "SALDO EM ABERTO POR FORNECEDOR POR PERIODO"].index[0]
novo_cabecalho = df.iloc[indice_conta]
df.columns = novo_cabecalho

lista_index = [0,1,2,3,4,5,6,7,8,9]
df = df.drop(lista_index, axis='index')


nome_arquivo = "Teste2"

df.to_excel(f"{caminho}\{nome_arquivo}.xlsx",index=True)

workbook = op.load_workbook(f"{caminho}\{nome_arquivo}.xlsx")

sheet = workbook.active