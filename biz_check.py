import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import datetime

def update_business_status(api_key, excel_file_path, progress_bar, root):
    df = pd.read_excel(excel_file_path)
    start_time = datetime.datetime.now()
    
    # 결과를 저장할 새로운 열을 추가합니다.
    df['사업자 상태'] = ''
    df['세금 유형'] = ''
    df['세금 유형 변경일'] = ''
    
    api_url = "https://api.odcloud.kr/api/nts-businessman/v1/status?serviceKey=" + api_key
    total_rows = len(df)
    
    # 100개씩 나누어 API 요청을 보냅니다.
    for start_idx in range(0, total_rows, 100):
        end_idx = min(start_idx + 100, total_rows)
        business_numbers = df['사업자번호'][start_idx:end_idx].tolist()
        data = {"b_no": [str(num) for num in business_numbers]}
        
        response = requests.post(api_url, json=data, headers={'Content-Type': 'application/json'}, verify=False)
        if response.status_code == 200:
            results = response.json()
            for result in results['data']:
                index = df[df['사업자번호'] == int(result['b_no'])].index[0]
                df.at[index, '사업자 상태'] = result.get('b_stt', 'No Data')
                df.at[index, '세금 유형'] = result.get('tax_type', 'No Data')
                df.at[index, '세금 유형 변경일'] = result.get('tax_type_change_dt', 'No Data')
        else:
            for index in range(start_idx, end_idx):
                df.at[index, '사업자 상태'] = 'Error'
                df.at[index, '세금 유형'] = 'Error'
                df.at[index, '세금 유형 변경일'] = 'Error'
        
        progress_value = int((end_idx + 1) / total_rows * 100)
        progress_bar['value'] = progress_value
        root.update_idletasks()
    
    df.to_excel(excel_file_path, index=False)
    end_time = datetime.datetime.now()
    time_diff = end_time - start_time
    
    messagebox.showinfo("완료", f"사업자 상태 업데이트가 완료되었습니다.\n소요 시간: {time_diff}")
    root.destroy()


def upload_file(api_key, progress_bar, root):
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            total_businesses = len(pd.read_excel(file_path))
            proceed = messagebox.askyesno("확인", f"총 {total_businesses}개의 사업자번호를 확인합니다. 계속하시겠습니까?")
            if proceed:
                update_business_status(api_key, file_path, progress_bar, root)
        except Exception as e:
            messagebox.showerror("오류", str(e))

def main(root):
    api_key = 'hoRuQGqHatZNJVYlmOeRK1H10ejjrHRPkwddmbLJtecpyFjxV4ObhOSZsMROb11eldnnNDJIiP1QY%2B0SvUZlJg%3D%3D'  # API 키
    instruction_label = tk.Label(root, text="컬럼명은 무조건 '사업자번호'로 기입해주세요.")
    instruction_label.pack(pady=10)
    upload_button = tk.Button(root, text="엑셀 파일 업로드", command=lambda: upload_file(api_key, progress_bar, root))
    upload_button.pack(pady=10)
    progress_bar = ttk.Progressbar(root, orient='horizontal', length=300, mode='determinate')
    progress_bar.pack(pady=10)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("사업자 상태 업데이트 프로그램")
    main(root)
    root.mainloop()
