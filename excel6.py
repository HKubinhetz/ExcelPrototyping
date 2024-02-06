# -------------------------------------- Exercises --------------------------------------
# 6) Manipulate an opened sheet
# Imports
import xlwings as xw
import pandas as pd
import numpy as np
from datetime import datetime

wb = xw.Book("example.xlsx")  # Workbook object
base = wb.sheets['Base']      # Worksheet object
# base["N3"].value = "Teste"    # Write on a cell

# Printing Cells
# print(base["C6"].value)

# Next part = Find indexes.

# print(base["C6"].value)     # Printing Cells
# base["N3"].value = "Teste"  # Write on a cell

# Next part = Find indexes
# Think I'll have to use pandas for that.
# Then use  <sheet.range((row,col)).value = "TEXT">

# So, off we go:
# Dataframe creation:
base_df = pd.DataFrame(base.used_range.value)  # Had to look that up
new_header = base_df.iloc[0]  # Grab the first row for the header
base_df = base_df[1:]  # Take the data less the header row
base_df.columns = new_header  # Set the header row as the df header

# Dataframe setup:
base_df = base_df.astype({'INSTALAÇÃO': np.int64})  # Setting column to int
col_index = base_df.columns.tolist().index('Chamado')


def writeticket():
    # TODO - Create safeties:
    # TODO - 1) Format edited Cells

    cdie = 691969
    row_index = int(base_df.index[base_df['INSTALAÇÃO'] == cdie][0])
    todaydate = datetime.today().strftime('%d/%m')
    data = todaydate + " - " + "TICKET00001"
    # cdie = int(input("Por favor informe a instalação que deseja abrir um chamado: "))
    # ticket = str(input("Por favor informe o número do chamado: "))

    cell = base[row_index, col_index]

    if cell.value is None:
        cell.value = data

    else:
        cell.value = data + "\n" + cell.value


writeticket()


