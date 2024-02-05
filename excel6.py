# -------------------------------------- Exercises --------------------------------------
# 6) Manipulate an opened sheet
# Imports
import xlwings as xw  # xlwings for Excel manipulation
import pandas as pd
import numpy as np

wb = xw.Book("example.xlsx")  # Workbook object
base = wb.sheets['Base']  # Worksheet object

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
base_df = base_df.set_index('INSTALAÇÃO')  # Setting a better index (unique, recognizable value)

# Testing
# print(base_df.head())
print(base_df["Chamado"].loc[737206])  # This is a great way of finding column and client!

# TODO - Now, it is needed to find the index of this cell


