import pandas as pd
import os
import shutil

# 엑셀 파일 경로
excel_path = './Test_Minseo/Test.xls'
# Before, After 폴더 경로
before_root = './Test_Minseo/Before'
after_root = './Test_Minseo/After'

# 엑셀 데이터 읽기
df = pd.read_excel(excel_path, sheet_name=0)

# PDF 파일명과 논문번호 매칭을 위한 열 이름 지정
# (이전) pdf_file 열: 'pdf_file', (변경)논문번호 열: '논문번호'
# 실제 열 이름에 맞게 수정 필요
pdf_file_col = '(이전)pdf_file'  # (이전) pdf 파일명 열
paper_num_col = '(변경)논문번호(N072'  # (변경) 논문번호 열
year_col = '발행년도'
vol_col = '권'
issue_col = '호'

# 모든 하위 폴더 순회
def process_pdfs():
    for folder_name in os.listdir(before_root):
        folder_path = os.path.join(before_root, folder_name)
        if not os.path.isdir(folder_path):
            continue
        for file_name in os.listdir(folder_path):
            if not file_name.lower().endswith('.pdf'):
                continue
            # 엑셀에서 해당 PDF 파일명 찾기
            row = df[df[pdf_file_col] == file_name]
            if row.empty:
                print(f"엑셀에서 {file_name} 정보를 찾을 수 없습니다.")
                continue
            row = row.iloc[0]
            new_pdf_name = str(row[paper_num_col]) + '.pdf'
            # 폴더명 생성
            year = str(int(row[year_col]))
            vol = f"{int(row[vol_col]):03d}"
            issue = f"{int(row[issue_col]):02d}"
            folder1 = f"{year}-{vol}-{issue}"
            folder2 = str(row[paper_num_col])
            dest_dir = os.path.join(after_root, folder1, folder2)
            os.makedirs(dest_dir, exist_ok=True)
            src_pdf = os.path.join(folder_path, file_name)
            dest_pdf = os.path.join(dest_dir, new_pdf_name)
            shutil.copy2(src_pdf, dest_pdf)
            print(f"{src_pdf} -> {dest_pdf}")

if __name__ == '__main__':
    process_pdfs()
