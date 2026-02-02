import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox

# PDF 이름 변경 및 이동 함수
def process_pdfs(excel_path, before_root, after_root, log_callback=None):
    try:
        df = pd.read_excel(excel_path, sheet_name=0)
        pdf_file_col = '(이전)pdf_file'
        paper_num_col = '(변경)논문번호(N072'
        year_col = '발행년도'
        vol_col = '권'
        issue_col = '호'
        count = 0
        for folder_name in os.listdir(before_root):
            folder_path = os.path.join(before_root, folder_name)
            if not os.path.isdir(folder_path):
                continue
            for file_name in os.listdir(folder_path):
                if not file_name.lower().endswith('.pdf'):
                    continue
                row = df[df[pdf_file_col] == file_name]
                if row.empty:
                    msg = f"엑셀에서 {file_name} 정보를 찾을 수 없습니다."
                    if log_callback: log_callback(msg)
                    continue
                row = row.iloc[0]
                new_pdf_name = str(row[paper_num_col]) + '.pdf'
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
                msg = f"{src_pdf} -> {dest_pdf}"
                if log_callback: log_callback(msg)
                count += 1
        if log_callback:
            log_callback(f"총 {count}개의 PDF가 이동/이름 변경되었습니다.")
        else:
            print(f"총 {count}개의 PDF가 이동/이름 변경되었습니다.")
    except Exception as e:
        if log_callback:
            log_callback(f"오류 발생: {e}")
        else:
            print(f"오류 발생: {e}")

# GUI 구현
def run_gui():
    root = tk.Tk()
    root.title('PDF 이름 변경 및 이동 자동화')
    root.geometry('600x400')

    excel_path = tk.StringVar()
    before_path = tk.StringVar()
    after_path = tk.StringVar()

    def select_excel():
        path = filedialog.askopenfilename(filetypes=[('Excel files', '*.xls;*.xlsx')])
        if path:
            excel_path.set(path)

    def select_before():
        path = filedialog.askdirectory()
        if path:
            before_path.set(path)

    def select_after():
        path = filedialog.askdirectory()
        if path:
            after_path.set(path)

    log_text = tk.Text(root, height=15, width=70)
    log_text.pack(pady=10)

    def log_callback(msg):
        log_text.insert(tk.END, msg + '\n')
        log_text.see(tk.END)
        root.update()

    def run_process():
        log_text.delete(1.0, tk.END)
        if not excel_path.get() or not before_path.get() or not after_path.get():
            messagebox.showwarning('입력 오류', '엑셀, Before, After 경로를 모두 선택하세요.')
            return
        process_pdfs(excel_path.get(), before_path.get(), after_path.get(), log_callback)
        messagebox.showinfo('완료', '작업이 완료되었습니다!')

    tk.Button(root, text='엑셀 파일 선택', command=select_excel).pack()
    tk.Entry(root, textvariable=excel_path, width=80).pack()
    tk.Button(root, text='Before 폴더 선택', command=select_before).pack()
    tk.Entry(root, textvariable=before_path, width=80).pack()
    tk.Button(root, text='After 폴더 선택', command=select_after).pack()
    tk.Entry(root, textvariable=after_path, width=80).pack()
    tk.Button(root, text='실행', command=run_process, bg='skyblue').pack(pady=10)

    root.mainloop()

if __name__ == '__main__':
    run_gui()
