import pandas as pd

import os

current_directory = os.getcwd()
print(os.getcwd())

# Load an Excel file
df = pd.read_excel("/path/to/your/excel/file.xlsx")

# Display the first few rows of the dataframe
print(df.head())