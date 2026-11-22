import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

# 全域變數：用來存放目前選擇的資料夾路徑，以及每個檔案對應的勾選狀態
current_folder = ""
file_vars = {}  # 格式：{"檔名.csv": tk.BooleanVar()}


def select_folder():
    """選擇資料夾並讀取裡面所有的 CSV 檔案"""
    global current_folder
    folder = filedialog.askdirectory(title="請選擇包含 CSV 檔案的資料夾")
    if not folder:
        return

    current_folder = folder
    folder_label.config(text=f"目前目錄：{current_folder}", fg="#1F4E78")
    load_csv_list()


def load_csv_list():
    """清空舊清單，重新掃描資料夾內的 CSV 並建立勾選框"""
    # 清空畫面上的舊勾選框
    for widget in list_frame.winfo_children():
        widget.destroy()
    file_vars.clear()

    # 找出所有 .csv 檔案並排序
    all_files = sorted(os.listdir(current_folder))
    csv_files = [f for f in all_files if f.lower().endswith(".csv")]

    if not csv_files:
        tk.Label(list_frame, text="此資料夾內找不到任何 CSV 檔案", fg="red", bg="white").pack(pady=20)
        return

    # 為每個 CSV 建立一個 Checkbutton
    for filename in csv_files:
        var = tk.BooleanVar(value=True)  # 預設全部勾選，方便操作
        file_vars[filename] = var
        chk = tk.Checkbutton(
            list_frame,
            text=filename,
            variable=var,
            anchor="w",
            bg="white",
            font=("微軟正黑體", 10)
        )
        chk.pack(fill="x", padx=10, pady=2)


def toggle_all(state):
    """全選或取消全選"""
    for var in file_vars.values():
        var.set(state)


def convert_single_csv(csv_path, excel_path):
    """核心轉換函式：讀取 CSV -> 轉存 Excel -> 美化標題與自動調整欄寬"""
    # 1. 讀取 CSV（dtype=str 確保店號、電話開頭的 0 不會消失）
    try:
        df = pd.read_csv(csv_path, encoding="utf-8-sig", dtype=str)
    except UnicodeDecodeError:
        df = pd.read_csv(csv_path, encoding="cp950", dtype=str)

    # 2. 寫入 Excel
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="資料明細")
        worksheet = writer.sheets["資料明細"]

        # 凍結首列（標題列）
        worksheet.freeze_panes = "A2"

        # 標題列樣式：深藍底白字加粗
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, name="微軟正黑體")

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")

        # 3. 自動適配欄位寬度
        for col in worksheet.columns:
            max_length = 0
            column_letter = col[0].column_letter  # 取得欄位代號如 A, B, C
            for cell in col:
                val = str(cell.value or "")
                # 中文字寬度算 2，英數字算 1
                length = sum(2 if ord(c) > 127 else 1 for c in val)
                if length > max_length:
                    max_length = length
            # 設定欄寬（加上 4 個字元的緩衝空間，避免太擁擠）
            worksheet.column_dimensions[column_letter].width = max_length + 4


def start_convert():
    """執行批次轉換"""
    if not current_folder:
        messagebox.showwarning("提醒", "請先選擇 CSV 資料夾！")
        return

    # 找出有被勾選的檔案
    selected_files = [name for name, var in file_vars.items() if var.get()]

    if not selected_files:
        messagebox.showwarning("提醒", "請至少勾選一個要轉換的 CSV 檔案！")
        return

    success_count = 0
    error_messages = []

    for filename in selected_files:
        csv_path = os.path.join(current_folder, filename)
        # 將副檔名 .csv 換成 .xlsx，存放在同一個資料夾下
        base_name = os.path.splitext(filename)[0]
        excel_path = os.path.join(current_folder, f"{base_name}.xlsx")

        try:
            convert_single_csv(csv_path, excel_path)
            success_count += 1
        except Exception as e:
            error_messages.append(f"{filename} 失敗原因：{e}")

    # 顯示完成結果
    if error_messages:
        err_text = "\n".join(error_messages)
        messagebox.showwarning("部分完成", f"成功轉換 {success_count} 個檔案。\n\n以下檔案發生錯誤：\n{err_text}")
    else:
        messagebox.showinfo("轉換完成", f"太棒了！已成功將 {success_count} 個檔案轉換為 Excel。\n檔案已儲存在原資料夾中。")


def on_mousewheel(event):
    """讓清單區塊支援滑鼠滾輪捲動"""
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# ==========================================
# 主視窗介面設定
# ==========================================
root = tk.Tk()
root.title("CSV 批次轉 Excel 小工具（自動欄寬版）")
root.geometry("520x480")
root.resizable(False, False)

# 1. 頂部：選擇資料夾區塊
top_frame = tk.Frame(root, pady=10, padx=15)
top_frame.pack(fill="x")

btn_select = tk.Button(
    top_frame,
    text="📂 選擇 CSV 資料夾",
    command=select_folder,
    font=("微軟正黑體", 10, "bold"),
    bg="#E8F0FE",
    padx=10,
    pady=4
)
btn_select.pack(side="left")

folder_label = tk.Label(top_frame, text="尚未選擇資料夾", fg="gray", font=("微軟正黑體", 9))
folder_label.pack(side="left", padx=10)

# 2. 中間：全選 / 取消全選 控制列
ctrl_frame = tk.Frame(root, padx=15, pady=5)
ctrl_frame.pack(fill="x")

tk.Button(ctrl_frame, text="全選", command=lambda: toggle_all(True), font=("微軟正黑體", 9), width=8).pack(side="left", padx=2)
tk.Button(ctrl_frame, text="取消全選", command=lambda: toggle_all(False), font=("微軟正黑體", 9), width=8).pack(side="left", padx=2)
tk.Label(ctrl_frame, text="（轉檔後會直接存於原目錄下）", fg="#666666", font=("微軟正黑體", 9)).pack(side="right")

# 3. 清單顯示區（加上 Canvas 與 Scrollbar 實現可捲動清單）
container = tk.Frame(root, padx=15)
container.pack(fill="both", expand=True)

canvas = tk.Canvas(container, bg="white", highlightthickness=1, highlightbackground="#CCCCCC")
scrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
list_frame = tk.Frame(canvas, bg="white")

list_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas.create_window((0, 0), window=list_frame, anchor="nw")
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind_all("<MouseWheel>", on_mousewheel)

canvas.pack(side="left", fill="both", expand=True)
scrollbar.pack(side="right", fill="y")

# 4. 底部：一鍵轉換按鈕
bottom_frame = tk.Frame(root, pady=15)
bottom_frame.pack(fill="x")

btn_convert = tk.Button(
    bottom_frame,
    text="⚡ 一鍵轉換為 Excel",
    command=start_convert,
    font=("微軟正黑體", 12, "bold"),
    bg="#1F4E78",
    fg="white",
    padx=25,
    pady=6
)
btn_convert.pack()

# 啟動視窗
root.mainloop()