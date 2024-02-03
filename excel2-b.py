# -------------------------------------- Exercises --------------------------------------
# 2) Read several specific cells from an existing Excel file and store them on variables;
# b) Now do it with Pandas!

# Imports
import pandas as pd

# Read Excel Sheet
df = pd.read_excel(".\\examplesheet\\example.xlsx")

# Iterating through rows
for index, row in df.iterrows():
    print(f"{row['INSTALAÇÃO']} - {row['NOME/SOBRENOME']}:"
          f" Possui tecnologia {row['Tecnologia']}")
