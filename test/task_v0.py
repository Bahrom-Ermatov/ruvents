# Import pandas
import pandas as pd

# Assign spreadsheet filename to `file`
file = 'task_support.xlsx'

# Load spreadsheet
xl = pd.ExcelFile(file)

# Print the sheet names
print(xl.sheet_names)

# Load a sheet into a DataFrame by name: df1
df1 = xl.parse('Tasks')

#print(df1.num1)

data = list(df1.num1)

print(data)



















