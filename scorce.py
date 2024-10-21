import requests
from bs4 import BeautifulSoup
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import csv

# 定義爬蟲函數
def scrape_grades():
    # 取得使用者輸入的資料
    username = username_entry.get()
    password = password_entry.get()
    grades_url = url_entry.get()
    save_path = save_path_entry.get()  # 取得存檔位置
    file_name = file_name_entry.get()   # 取得檔案名稱
    
    # 檢查存檔位置和檔案名稱是否為空
    if not save_path:
        messagebox.showerror("錯誤", "請選擇存檔位置。")
        return
    if not file_name:
        messagebox.showerror("錯誤", "請輸入檔案名稱。")
        return
    
    # 檢查檔案名稱是否以.csv結尾
    if not file_name.endswith('.csv'):
        file_name += '.csv'  # 如果沒有以.csv結尾，自動添加
    
    # 建立Session物件
    session = requests.Session()
    
    login_page = session.get("https://moodle.nhu.edu.tw/login/index.php")

    soup = BeautifulSoup(login_page.content, 'html.parser')
    logintoken = soup.find('input', {'name': 'logintoken'})['value']

    login_data = {
        'username': username,  # 使用者輸入的帳號
        'password': password,  # 使用者輸入的密碼
        'logintoken': logintoken       # 使用獲取的logintoken
    }

    response = session.post("https://moodle.nhu.edu.tw/login/index.php", data=login_data)

    if "登出" in response.text:
        print("登入成功！")
    else:
        print("登入失敗，請檢查帳號和密碼。")
        messagebox.showerror("登入失敗", "請檢查帳號和密碼。")
        return
    
    # 擷取成績頁面的內容
    grades_response = session.get(grades_url)
    
    # 使用BeautifulSoup解析HTML
    soup = BeautifulSoup(grades_response.text, 'html.parser')
    
    # 找到所有包含學號、姓名和分數的行
    rows = soup.find_all('tr')  # 調整class名稱以匹配你的需求
    result = []
    
    for row in rows:
        # 提取學號和姓名
        name_cell = row.find('td', class_='cell c2')
        student_id_cell = row.find('label')
        
        if name_cell and student_id_cell:
            name = name_cell.get_text(strip=True)
            student_id = student_id_cell.get_text(strip=True).replace('選擇', '')
            student_data = student_id.split()
            student_id = student_data[1]
            name = student_data[0]
            # 提取分數
            score_cell = row.find('td', class_='cell c12')
            if score_cell:
                score = score_cell.get_text(strip=True)
                
                # 將結果添加到列表中
                result.append([student_id, name, score])
    
    # 顯示結果到文本框
    result_text.delete(1.0, tk.END)  # 清空文本框
    if result:
        for student in result:
            result_text.insert(tk.END, f"學號: {student[0]}, 姓名: {student[1]}, 分數: {student[2]}\n")
        
        # 匯出成績到CSV
        export_to_csv(result, save_path, file_name)
    else:
        result_text.insert(tk.END, "沒有找到任何成績。")

def export_to_csv(data, save_path, file_name):
    # 指定CSV文件名
    filename = f"{save_path}/{file_name}"  # 使用選擇的路徑和檔案名稱
    # 寫入CSV文件
    with open(filename, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        # 寫入標題行
        writer.writerow(['學號', '姓名', '分數'])
        # 寫入數據行
        writer.writerows(data)
    
    messagebox.showinfo("匯出成功", f"成績已成功匯出到 {filename}")

def select_save_path():
    # 打開文件對話框選擇存檔位置
    path = filedialog.askdirectory()  # 選擇目錄
    if path:
        save_path_entry.delete(0, tk.END)  # 清空輸入框
        save_path_entry.insert(0, path)     # 設置選擇的路徑

# 建立主視窗
root = tk.Tk()
root.title("Moodle 成績抓取工具")

# 使用者介面
tk.Label(root, text="帳號:").grid(row=0, column=0, padx=10, pady=10)
username_entry = tk.Entry(root)
username_entry.grid(row=0, column=1, padx=10, pady=10)

tk.Label(root, text="密碼:").grid(row=1, column=0, padx=10, pady=10)
password_entry = tk.Entry(root, show='*')
password_entry.grid(row=1, column=1, padx=10, pady=10)

tk.Label(root, text="成績頁面網址:").grid(row=2, column=0, padx=10, pady=10)
url_entry = tk.Entry(root, width=50)
url_entry.grid(row=2, column=1, padx=10, pady=10)

tk.Label(root, text="存檔位置:").grid(row=3, column=0, padx=10, pady=10)
save_path_entry = tk.Entry(root, width=50)
save_path_entry.grid(row=3, column=1, padx=10, pady=10)

tk.Label(root, text="檔案名稱:").grid(row=4, column=0, padx=10, pady=10)
file_name_entry = tk.Entry(root, width=50)
file_name_entry.grid(row=4, column=1, padx=10, pady=10)

# 按鈕
select_path_button = tk.Button(root, text="選擇存檔位置", command=select_save_path)
select_path_button.grid(row=3, column=2, padx=10, pady=10)

scrape_button = tk.Button(root, text="抓取成績", command=scrape_grades)
scrape_button.grid(row=5, column=0, columnspan=3, padx=10, pady=10)

# 顯示結果的文本框
result_text = scrolledtext.ScrolledText(root, width=100, height=20)
result_text.grid(row=6, column=0, columnspan=3, padx=10, pady=10)

# 開始主循環
root.mainloop()
