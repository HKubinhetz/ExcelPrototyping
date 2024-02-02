# -------------------------------------- Exercises --------------------------------------
# 1) Create a new Excel file and write it on a specific cell;

# Imports
from openpyxl import Workbook


# Exercise declaration
def excelprototyping1():

    # Creating the workbook object
    wb = Workbook()

    # Finding the active worksheet
    ws = wb.active
    ws['A1'] = "Hi mom!"

    # Appending an entire row
    ws.append([1, 2, 3])

    # Saving the workbook
    wb.save("excel1.xlsx")


# Exercise execution
if __name__ == '__main__':
    excelprototyping1()



