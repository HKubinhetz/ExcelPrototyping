# -------------------------------------- Exercises --------------------------------------
# 4) Find the index of a cell
# 5) Select and read a cell from a given index
# My feeling is we ought start by using pandas right away
# Update: my feeling was right! :)

# Imports
import pandas as pd
import numpy as np

# Read Excel Sheet
df = pd.read_excel(".\\examplesheet\\example.xlsx")

# The following columns will compose our desired indexes.
columnlist = ['INSTALAÇÃO', 'NOME/SOBRENOME', 'Lacunas']
indexlist = []

# Finding the index of a column - columns.get_loc function.
# Iterating through those and building a list.
for columnlabel in columnlist:
    found_index = df.columns.get_loc(columnlabel)
    indexlist.append(found_index)

# Finding the index of a row
row_index = np.where(df['NOME/SOBRENOME'] == "AGC VIDROS DO BRASIL LTDA.")[0][0]

# Creating a list with the desired information:
datalist = []
for i in range(len(indexlist)):
    info = df.iloc[row_index, indexlist[i]]
    datalist.append(info)

# Printing the final result
print(f"O cliente {datalist[0]} - {datalist[1]} possui {datalist[2]} lacunas;")

