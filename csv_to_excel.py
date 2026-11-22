import os
import sys
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment

print("程式開始啟動...")  # 觀察終端機有沒有印出這行

current_folder = ""
file_vars = {}

THEME = {
    "bg_main": "#F4F6F9",
    "bg_card": "#FFFFFF",
    "header_bg": "#1E293B",
    "header_fg": "#FFFFFF",
    "primary": "#2563EB",
    "primary_hover": "#1D4ED8",
    "teal": "#0D9488",
    "coral": "#F97316",
    "merge_bg": "#EEF2FF",
    "merge_fg": "#4338CA",
    "convert": "#10B981",
    "convert_hover": "#059669",
    "text_dark": "#1E293B",
    "text_muted": "#64748B",
    "border": "#E2E8F0"
}


def get_output_dir():
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))

    output_dir = os.path.join(base_dir, "output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def update_count_label():
    total = len(file_vars)
    checked = sum(1 for var in file_vars.values() if var.get())
    if total == 0:
        count_label.config(text="尚未載入檔案", fg=THEME["text_muted"])
    else:
        count_label.config(text=f"共 {total} 個檔案 ｜ 已勾選 {checked} 個", fg=THEME["primary"])


def select_folder():
    global current_folder
    folder = filedialog.askdirectory(title="請選擇包含 CSV 檔案的資料夾")
    if not folder:
        return

    current_folder = folder
    short_path = folder if len(folder) <= 45 else "..." + folder[-42:]
    folder_label.config(text=f"📁 {short_path}", fg=THEME["text_dark"])
    load_csv_list()


def load_csv_list():
    for widget in list_frame.winfo_children():
        widget.destroy()
    file_vars.clear()

    all_files = sorted(os.listdir(current_folder))
    csv_files = [f for f in all_files if f.lower().endswith(".csv")]

    if not csv_files:
        tk.Label(
            list_frame,
            text="⚠️ 此資料夾內找不到任何 .csv 檔案",
            fg="#EF4444",
            bg=THEME["bg_card"],
            font=("微軟正黑體", 10, "bold")
        ).pack(pady=40)
        update_count_label()
        return

    for idx, filename in enumerate(csv_files):
        var = tk.BooleanVar(value=True)
        file_vars[filename] = var

        row_bg = "#FFFFFF" if idx % 2 == 0 else "#F8FAFC"

        chk = tk.Checkbutton(
            list_frame,
            text=f"  📄 {filename}",
            variable=var,
            command=update_count_label,
            anchor="w",
            bg=row_bg,
            activebackground="#EFF6FF",
            selectcolor="#FFFFFF",
            fg=THEME["text_dark"],
            font=("微軟正黑體", 10),
            padx=10,
            pady=4,
            cursor="hand2"
        )
        chk.pack(fill="x", padx=4, pady=1)

    update_count_label()


def toggle_all(state):
    for var in file_vars.values():
        var.set(state)
    update_count_label()


def read_csv_safely(csv_path):
    try:
        return pd.read_csv(csv_path, encoding="utf-8-sig", dtype=str)
    except UnicodeDecodeError:
        return pd.read_csv(csv_path, encoding="cp950", dtype=str)


def style_worksheet(worksheet):
    worksheet.freeze_panes = "A2"

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, name="微軟正黑體")

    for cell in worksheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in worksheet.columns:
        max_length = 0
        column_letter = col[0].column_letter
        for cell in col:
            val = str(cell.value or "")
            length = sum(2 if ord(c) > 127 else 1 for c in val)
            if length > max_length:
                max_length = length
        worksheet.column_dimensions[column_letter].width = max_length + 4


def make_safe_sheet_name(raw_name, used_names):
    for ch in ['\\', '/', '?', '*', '[', ']', ':']:
        raw_name = raw_name.replace(ch, '_')

    name = raw_name[:31] if raw_name else "Sheet"

    original = name
    counter = 1
    while name in used_names:
        suffix = f"_{counter}"
        name = original[:31 - len(suffix)] + suffix
        counter += 1

    used_names.add(name)
    return name


def start_convert():
    if not current_folder:
        messagebox.showwarning("提醒", "請先選擇 CSV 資料夾！")
        return

    selected_files = [name for name, var in file_vars.items() if var.get()]
    if not selected_files:
        messagebox.showwarning("提醒", "請至少勾選一個要轉換的 CSV 檔案！")
        return

    output_dir = get_output_dir()
    is_merge = merge_var.get()

    if is_merge:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"合併報表_{timestamp}.xlsx"
        excel_path = os.path.join(output_dir, excel_filename)

        used_sheet_names = set()
        success_count = 0
        error_messages = []

        try:
            with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                for filename in selected_files:
                    csv_path = os.path.join(current_folder, filename)
                    base_name = os.path.splitext(filename)[0]
                    sheet_name = make_safe_sheet_name(base_name, used_sheet_names)

                    try:
                        df = read_csv_safely(csv_path)
                        df.to_excel(writer, index=False, sheet_name=sheet_name)
                        style_worksheet(writer.sheets[sheet_name])
                        success_count += 1
                    except Exception as e:
                        error_messages.append(f"{filename} 讀取失敗：{e}")

            if success_count == 0:
                if os.path.exists(excel_path):
                    os.remove(excel_path)
                messagebox.showerror("轉換失敗", "所有檔案皆無法讀取，未產生 Excel 檔案。")
                return

            msg = f"已將 {success_count} 個檔案合併至：\noutput\\{excel_filename}"
            if error_messages:
                msg += "\n\n以下檔案略過：\n" + "\n".join(error_messages)
                messagebox.showwarning("部分完成", msg)
            else:
                messagebox.showinfo("合併完成", msg)

            os.startfile(output_dir)

        except Exception as e:
            messagebox.showerror("錯誤", f"建立合併檔案時發生錯誤：\n{e}")

    else:
        success_count = 0
        error_messages = []

        for filename in selected_files:
            csv_path = os.path.join(current_folder, filename)
            base_name = os.path.splitext(filename)[0]
            excel_path = os.path.join(output_dir, f"{base_name}.xlsx")

            try:
                df = read_csv_safely(csv_path)
                with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="資料明細")
                    style_worksheet(writer.sheets["資料明細"])
                success_count += 1
            except Exception as e:
                error_messages.append(f"{filename} 失敗原因：{e}")

        if error_messages:
            err_text = "\n".join(error_messages)
            messagebox.showwarning("部分完成", f"成功轉換 {success_count} 個檔案至 output 資料夾。\n\n錯誤明細：\n{err_text}")
        else:
            messagebox.showinfo("轉換完成", f"已成功將 {success_count} 個檔案轉換並儲存至：\n{output_dir}")

        if success_count > 0:
            os.startfile(output_dir)


def on_mousewheel(event):
    canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")


# ==========================================
# 主視窗介面設定
# ==========================================
root = tk.Tk()
root.title("CSV 批次轉 Excel 小工具")
root.geometry("600x580")
root.resizable(False, False)
root.configure(bg=THEME["bg_main"])

# 強制視窗開啟時浮在最上層，避免躲在編輯器後面
root.lift()
root.attributes("-topmost", True)
root.after(200, lambda: root.attributes("-topmost", False))

# 1. 頂部深色橫幅
header_frame = tk.Frame(root, bg=THEME["header_bg"], pady=15, padx=20)
header_frame.pack(fill="x")

tk.Label(
    header_frame,
    text="📊 CSV 批次轉 Excel 轉換器",
    bg=THEME["header_bg"],
    fg=THEME["header_fg"],
    font=("微軟正黑體", 14, "bold")
).pack(anchor="w")

tk.Label(
    header_frame,
    text="自動防亂碼 • 保留前導零 • 智慧調整欄寬 • 支援多工作表合併",
    bg=THEME["header_bg"],
    fg="#94A3B8",
    font=("微軟正黑體", 9)
).pack(anchor="w", pady=(2, 0))

# 2. 選擇資料夾卡片區塊
folder_card = tk.Frame(
    root,
    bg=THEME["bg_card"],
    highlightbackground=THEME["border"],
    highlightthickness=1,
    padx=15,
    pady=12
)
folder_card.pack(fill="x", padx=18, pady=(15, 10))

btn_select = tk.Button(
    folder_card,
    text="📂 選擇 CSV 資料夾",
    command=select_folder,
    font=("微軟正黑體", 10, "bold"),
    bg=THEME["primary"],
    fg="white",
    activebackground=THEME["primary_hover"],
    activeforeground="white",
    relief="flat",
    padx=14,
    pady=6,
    cursor="hand2"
)
btn_select.pack(side="left")

btn_select.bind("<Enter>", lambda e: btn_select.config(bg=THEME["primary_hover"]))
btn_select.bind("<Leave>", lambda e: btn_select.config(bg=THEME["primary"]))

folder_label = tk.Label(
    folder_card,
    text="尚未選擇資料夾，請點擊左側按鈕",
    bg=THEME["bg_card"],
    fg=THEME["text_muted"],
    font=("微軟正黑體", 9)
)
folder_label.pack(side="left", padx=12)

# 3. 控制列
ctrl_frame = tk.Frame(root, bg=THEME["bg_main"], padx=18, pady=4)
ctrl_frame.pack(fill="x")

btn_all = tk.Button(
    ctrl_frame,
    text="✔ 全選",
    command=lambda: toggle_all(True),
    font=("微軟正黑體", 9, "bold"),
    bg=THEME["teal"],
    fg="white",
    relief="flat",
    padx=10,
    pady=4,
    cursor="hand2"
)
btn_all.pack(side="left", padx=(0, 6))

btn_none = tk.Button(
    ctrl_frame,
    text="✖ 取消全選",
    command=lambda: toggle_all(False),
    font=("微軟正黑體", 9, "bold"),
    bg=THEME["coral"],
    fg="white",
    relief="flat",
    padx=10,
    pady=4,
    cursor="hand2"
)
btn_none.pack(side="left")

merge_box = tk.Frame(ctrl_frame, bg=THEME["merge_bg"], padx=8, pady=2)
merge_box.pack(side="right")

merge_var = tk.BooleanVar(value=True)
chk_merge = tk.Checkbutton(
    merge_box,
    text="📑 合併模式（同檔多工作表）",
    variable=merge_var,
    bg=THEME["merge_bg"],
    activebackground=THEME["merge_bg"],
    fg=THEME["merge_fg"],
    selectcolor="white",
    font=("微軟正黑體", 9, "bold"),
    cursor="hand2"
)
chk_merge.pack()

# 4. 清單卡片區塊
list_card = tk.Frame(
    root,
    bg=THEME["bg_card"],
    highlightbackground=THEME["border"],
    highlightthickness=1
)
list_card.pack(fill="both", expand=True, padx=18, pady=(6, 10))

status_bar = tk.Frame(list_card, bg="#F8FAFC", padx=12, pady=6)
status_bar.pack(fill="x")

tk.Label(
    status_bar,
    text="檔案清單",
    bg="#F8FAFC",
    fg=THEME["text_dark"],
    font=("微軟正黑體", 9, "bold")
).pack(side="left")

count_label = tk.Label(
    status_bar,
    text="尚未載入檔案",
    bg="#F8FAFC",
    fg=THEME["text_muted"],
    font=("微軟正黑體", 9, "bold")
)
count_label.pack(side="right")

tk.Frame(list_card, bg=THEME["border"], height=1).pack(fill="x")

canvas = tk.Canvas(list_card, bg=THEME["bg_card"], highlightthickness=0)
scrollbar = tk.Scrollbar(list_card, orient="vertical", command=canvas.yview)
list_frame = tk.Frame(canvas, bg=THEME["bg_card"])

list_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas_window = canvas.create_window((0, 0), window=list_frame, anchor="nw")
canvas.bind("<Configure>", lambda e: canvas.itemconfig(canvas_window, width=e.width))
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind_all("<MouseWheel>", on_mousewheel)

canvas.pack(side="left", fill="both", expand=True, pady=4)
scrollbar.pack(side="right", fill="y")

# 5. 底部執行區塊
bottom_frame = tk.Frame(root, bg=THEME["bg_main"], pady=10, padx=18)
bottom_frame.pack(fill="x")

btn_convert = tk.Button(
    bottom_frame,
    text="🚀 立即轉換為精美 Excel",
    command=start_convert,
    font=("微軟正黑體", 12, "bold"),
    bg=THEME["convert"],
    fg="white",
    activebackground=THEME["convert_hover"],
    activeforeground="white",
    relief="flat",
    pady=10,
    cursor="hand2"
)
btn_convert.pack(fill="x")

btn_convert.bind("<Enter>", lambda e: btn_convert.config(bg=THEME["convert_hover"]))
btn_convert.bind("<Leave>", lambda e: btn_convert.config(bg=THEME["convert"]))

tk.Label(
    bottom_frame,
    text="自動儲存至程式同層的 output 資料夾，完成後將自動開啟",
    bg=THEME["bg_main"],
    fg=THEME["text_muted"],
    font=("微軟正黑體", 8)
).pack(pady=(6, 0))

print("準備開啟視窗...")
root.mainloop()