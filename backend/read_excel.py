"""Leitura simples de um arquivo Excel para inspecionar celulas nao vazias."""

import openpyxl

# Abre a planilha padrao (ajuste o nome se necessario)
wb = openpyxl.load_workbook('Ficha de Manutenção Padrão Rev. - 05_Julho_2024.xlsx')
ws = wb.active

# Varre as primeiras linhas/colunas e imprime celulas com valor
for row_num in range(1, 35):
    for col_num in range(1, 6):
        cell = ws.cell(row_num, col_num)
        if cell.value:
            print(f"{cell.coordinate}: {cell.value}")
