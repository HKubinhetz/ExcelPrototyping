# -------------------------------------- Exercises --------------------------------------
# 6) Manipulate an opened sheet
# Imports
import xlwings as xw
import pandas as pd
import numpy as np

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
# row_index = int(base_df.index[base_df['INSTALAÇÃO'] == 691969][0]) + 1  # Adding one to exclude header
# print(base[row_index, col_index].value)


def itsm_ticket():
    # TODO - Create safeties:
    # TODO - 1) Check if there is something on that cell (if so, append);
    # TODO - 2) Check if the line is correct with pandas! A simple if loop should do the trick
    cdie = int(input("Por favor informe a instalação que deseja abrir um chamado: "))
    row = int(base_df.index[base_df['INSTALAÇÃO'] == cdie][0])
    ticket = str(input("Por favor informe o número do chamado: "))
    base[row, col_index].value = ticket


itsm_ticket()
