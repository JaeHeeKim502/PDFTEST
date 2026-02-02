import pandas as pd

# 엑셀 파일 경로
excel_path = './Test_Minseo/Test.xls'

# 첫 번째 시트의 헤더(열 이름)만 출력
df = pd.read_excel(excel_path, sheet_name=0)
print(df.columns.tolist())
