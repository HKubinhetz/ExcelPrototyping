# 9) Finally, save something at an Excel Sheet and run a VBA macro from there, that then runs another py script.
# Don't forget to pass on variables.

# Imports
import xlwings as xw
from pathlib import Path

# Creating Workbook
wb = xw.Book("simplesum.xlsm")

# Finding Macro
macro = wb.macro("simplesub")

# Running Macro
# macro(3, 3)


def runfromxl():
    print("This is running from Excel!")
    file_path = Path(__file__).parent / 'fromexcel.txt'

    with open(file_path, 'w') as f:
        f.write(f'This came from excel 2.0!')


# Then, write this on VBA (use pythonpath to point the correct script)

"""
Sub runfromvba()

    runpython ("import excel9; excel9.runfromxl()")

End Sub

"""
