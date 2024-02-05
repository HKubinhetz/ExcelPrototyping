# -------------------------------------- Exercises --------------------------------------
# 6) Manipulate an opened sheet
# Imports
import xlwings as xw          # Import
wb = xw.Book("example.xlsx")  # Workbook object
base = wb.sheets['Base']      # Worksheet object
base["N3"].value = "Teste"    # Write on a cell

# Printing Cells
# print(base["C6"].value)


# Next part = Find indexes.






