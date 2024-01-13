import pandas as pd
import requests
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

def update_business_status(api_key, excel_file_path, progress_bar, root):
    df = pd.read_excel(excel_file_path)
    
    # 필요한 새로운 열 추가
    df['사업자 상태'] = ''
    df['세금 유형'] = ''
    df['세금 유형 변경일'] = ''

    api_url = "https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=" + api_key
    total_rows = len(df)

    for index, row in df.iterrows():
        business_number = row['사업자번호']
        data = {"b_no": [str(business_number)]}
        response = requests.post(api_url, data=json.dumps(data), headers={'Content-Type': 'application/json'})

        if response.status_code == 200:
            result = response.json()
            if result['data']:
                data_entry = result['data'][0]
                df.at[index, '사업자 상태'] = data_entry.get('b_stt', 'No Data')
                df.at[index, '세금 유형'] = data_entry.get('tax_type', 'No Data')
                df.at[index, '세금 유형 변경일'] = data_entry.get('tax_type_change_dt', 'No Data')
            else:
                df.at[index, '사업자 상태'] = 'No Data'
                df.at[index, '세금 유형'] = 'No Data'
                df.at[index, '세금 유형 변경일'] = 'No Data'
        else:
            df.at[index, '사업자 상태'] = 'Error'
            df.at[index, '세금 유형'] = 'Error'
            df.at[index, '세금 유형 변경일'] = 'Error'

        progress_bar['value'] = (index + 1) / total_rows * 100
        root.update_idletasks()

    df.to_excel(excel_file_path, index=False)
    messagebox.showinfo("완료", "사업자 상태 업데이트가 완료되었습니다.")
    root.destroy()

def upload_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            total_businesses = len(pd.read_excel(file_path))
            proceed = messagebox.askyesno("확인", f"총 {total_businesses}개의 사업자번호를 확인합니다. 계속하시겠습니까?")
            if proceed:
                update_business_status(api_key, file_path, progress_bar, root)
        except Exception as e:
            messagebox.showerror("오류", str(e))

root = tk.Tk()
root.title("사업자 상태 업데이트 프로그램")

api_key = 'you_api_key'  # 여기에 실제 API 키를 입력하세요

instruction_label = tk.Label(root, text="컬럼명은 무조건 '사업자번호'로 기입해주세요.")
instruction_label.pack(pady=10)

upload_button = tk.Button(root, text="엑셀 파일 업로드", command=upload_file)
upload_button.pack(pady=10)

progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
progress_bar.pack(pady=10)

root.mainloop()
