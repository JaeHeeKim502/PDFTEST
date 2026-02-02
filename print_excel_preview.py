import pandas as pd

excel_path = './Test_Minseo/Test.xls'

df = pd.read_excel(excel_path, sheet_name=0)
print(df.head(3))
print(df.columns.tolist())
