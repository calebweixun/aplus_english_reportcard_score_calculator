import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import subprocess
import webbrowser

folder_path = os.path.join(os.path.expanduser("~"), "Desktop")


def process_files():
    # 讀取輸入的Excel檔案路徑
    excel_files = [entry1.get(), entry2.get(), entry3.get(), entry4.get()]

    # 檢查所有路徑是否有效
    for paths in excel_files:
        paths_now = excel_files.index(paths)
        if not os.path.isfile(paths):
            messagebox.showerror("錯誤", f"無效的文件路徑: Level {paths_now+1}")
            return

    tables = ['A', 'P', 'L', 'U', 'S']

    output_text.delete(1.0, tk.END)

    for tabname in tables:
        combined_data = pd.DataFrame()
        for_now = tables.index(tabname)
        window_text_show(f"開始處理分頁-{tabname}\n")

        for file in excel_files:

            window_text_show(f"處理中檔案：{file},,\n")

            try:
                excel_file = pd.ExcelFile(file)
                sheet_names = excel_file.sheet_names

                df = pd.read_excel(file, sheet_name=for_now)

                df.insert(len(df.columns), 'Tab', tabname)
                df.insert(len(df.columns), 'File', file)

                # output_text.insert(tk.END, f"\n{df.columns},,\n")
                # output_text.update()
                combined_data = pd.concat(
                    [combined_data, df], ignore_index=True)
                # output_text.insert(tk.END, f"\n{combined_data.columns},,\n")
                # output_text.update()
            except Exception as e:
                window_text_show(f"Error reading {file}: {str(e)}\n")
                continue

        window_text_show(f"分數計算中,,")
        for index, row in combined_data.iterrows():
            count_col = [4, 6, 8, 10, 12]
            countval = 0
            for col in count_col:
                col_val = str(row.iloc[col])
                if 'A+' in col_val:
                    countval += 1

            combined_data.loc[index, 'Score'] = countval

            # 去除姓名、英文名、班級欄位資料中的空格，避免分組錯誤
            combined_data.at[index, combined_data.columns[1]] = str(
                combined_data.at[index, combined_data.columns[1]]).strip()
            combined_data.at[index, combined_data.columns[2]] = str(
                combined_data.at[index, combined_data.columns[2]]).strip()
            combined_data.at[index, combined_data.columns[3]] = str(
                combined_data.at[index, combined_data.columns[3]]).strip()

        result = combined_data.groupby(['Ch.', 'Eng.', 'Class'])[
            'Score'].sum().reset_index()

        window_text_show(f",完成,,輸出資料中,")

        combined_data.to_excel(
            f'{folder_path}/raw_count_val_{tabname}.xlsx', index=False)
        result.to_excel(f'{folder_path}/final_sum_{tabname}.xlsx', index=False)

        window_text_show(f",完成\n\n")

    window_text_show(f"本次處理完成\n\n")


def open_save_folder():
    # 獲取當前應用程式的目錄
    # folder_path = os.getcwd()
    window_text_show(f"打開存檔資料夾 >>{folder_path}<< \n")

    subprocess.run(['open', folder_path])


def browser1():
    entry1.delete(0, tk.END)
    entry1.insert(0, filedialog.askopenfilename())


def browser2():
    entry2.delete(0, tk.END)
    entry2.insert(0, filedialog.askopenfilename())


def browser3():
    entry3.delete(0, tk.END)
    entry3.insert(0, filedialog.askopenfilename())


def browser4():
    entry4.delete(0, tk.END)
    entry4.insert(0, filedialog.askopenfilename())


def openGithub(event):
    webbrowser.open_new(
        "https://github.com/calebweixun/aplus_english_reportcard_score_calculator")


def center_window(root, width=300, height=200):
    # 获取屏幕宽度和高度
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 计算窗口的 x 和 y 坐标
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # 设置窗口的大小和位置
    root.geometry(f'{width}x{height}+{x}+{y}')


def window_text_show(text):
    output_text.insert(tk.END, text)
    output_text.update()
    output_text.yview_moveto(1.0)


root = tk.Tk()
root.title("APlus Dept.English A+ Score Calculator")
# root.iconbitmap('iconfile')
center_window(root, 745, 580)

root.resizable(False, False)


tk.Label(root, text="＊請依序選擇各 Level 的 Excel 檔案，程式將會自動處理資料\n＞＞輸出 raw_count_val 為合併校驗計算的結果，供驗證使用\n＞＞輸出 final_sum 則是每個人加總完之後的結果\n\n＊請特別注意各檔案的各表欄位標頭需相同，標頭若是不確定可以複製其中一檔案的標頭進行覆蓋\n＞若是欄位標頭有不同的情況，可能會出現資料錯誤的情形。", font=(
    'Helvetica', 16), anchor='w', justify='left').grid(row=1, column=0, columnspan=10, padx=23, pady=5, sticky='W')


tk.Label(root, text="Level 1:").grid(row=4, column=0, padx=5, pady=2)
entry1 = tk.Entry(root, width=55, justify='right')
entry1.grid(row=4, column=1, padx=5, pady=2)
tk.Button(root, text="Browse", command=browser1).grid(
    row=4, column=2, padx=5, pady=2, sticky='EW')

tk.Label(root, text="Level 2:").grid(row=5, column=0, padx=5, pady=2)
entry2 = tk.Entry(root, width=55, justify='right')
entry2.grid(row=5, column=1, padx=5, pady=2)
tk.Button(root, text="Browse", command=browser2).grid(
    row=5, column=2, padx=5, pady=2, sticky='EW')

tk.Label(root, text="Level 3:").grid(row=6, column=0, padx=5, pady=2)
entry3 = tk.Entry(root, width=55, justify='right')
entry3.grid(row=6, column=1, padx=5, pady=2)
tk.Button(root, text="Browse", command=browser3).grid(
    row=6, column=2, padx=5, pady=2, sticky='EW')

tk.Label(root, text="Level 4:").grid(
    row=7, column=0, padx=5, pady=2, sticky='EW')
entry4 = tk.Entry(root, width=55, justify='right')
entry4.grid(row=7, column=1, padx=5, pady=2)
tk.Button(root, text="Browse", command=browser4).grid(
    row=7, column=2, padx=5, pady=2, sticky='EW')

# 按鈕
tk.Button(root, text="開啟存檔資料夾", command=open_save_folder).grid(
    row=8, column=0, columnspan=1, sticky='EW')
tk.Button(root, text="開始處理檔案", command=process_files).grid(
    row=8, column=1, columnspan=2, sticky='EW')

# 顯示資訊的文字框
output_text = scrolledtext.ScrolledText(root, width=100, height=20)
output_text.grid(row=9, column=0, columnspan=3, padx=10, pady=1)


tk.Label(root, text="目前輸出路徑為 "+folder_path, font=(
    'Helvetica', 12), anchor='w', justify='left').grid(row=10, column=0, columnspan=10, padx=20, pady=0, sticky='W')
hyperlink = tk.Label(root, text="create by CalebZhang v240520", font=(
    'Helvetica', 12), anchor='w', justify='right', fg="#3E84DE", cursor="hand2")
hyperlink.grid(row=10, column=1, columnspan=10, padx=20, pady=0, sticky='E')

hyperlink.bind("<Button-1>", openGithub)

# 啟動主循環
root.mainloop()
