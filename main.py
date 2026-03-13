import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side, Font

ROWS = 5
COLS = 6
TOTAL = ROWS * COLS

names = []
seats = []

# ======================
# 读取Excel名单
# ======================

def load_excel():

    global names

    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, header=None)
        names = df.iloc[:, 0].dropna().tolist()

        status_label.config(text=f"已导入 {len(names)} 名学生")

    except Exception as e:
        messagebox.showerror("读取失败", str(e))


# ======================
# 生成随机座位
# ======================

def generate_seats():

    global seats

    if len(names) == 0:
        messagebox.showwarning("提示", "请先导入名单")
        return

    if len(names) > TOTAL:
        messagebox.showwarning("提示", f"人数超过 {TOTAL}")
        return

    shuffled = names[:]
    random.shuffle(shuffled)

    while len(shuffled) < TOTAL:
        shuffled.append("")

    seats = [shuffled[i*COLS:(i+1)*COLS] for i in range(ROWS)]

    draw_seats()


# ======================
# 绘制教室布局
# ======================

def draw_seats():

    canvas.delete("all")

    cell_w = 110
    cell_h = 50

    start_x = 40
    start_y = 100

    # 讲台
    canvas.create_rectangle(
        start_x,
        30,
        start_x + COLS * cell_w,
        70,
        fill="#dddddd"
    )

    canvas.create_text(
        start_x + COLS * cell_w / 2,
        50,
        text="讲台",
        font=("Arial", 14, "bold")
    )

    for r in range(ROWS):
        for c in range(COLS):

            x1 = start_x + c * cell_w
            y1 = start_y + r * cell_h
            x2 = x1 + cell_w
            y2 = y1 + cell_h

            canvas.create_rectangle(x1, y1, x2, y2)

            name = seats[r][c]

            canvas.create_text(
                (x1+x2)/2,
                (y1+y2)/2,
                text=name,
                font=("Arial", 10)
            )


# ======================
# 导出Excel座位表
# ======================

def export_excel():

    if not seats:
        messagebox.showwarning("提示", "请先生成座位")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx"
    )

    if not save_path:
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "座位表"

    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin")
    )

    align = Alignment(horizontal="center", vertical="center")

    # 写入讲台
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=COLS)
    ws.cell(1,1).value = "讲台"
    ws.cell(1,1).alignment = align
    ws.cell(1,1).font = Font(size=14, bold=True)

    # 写入座位
    for r in range(ROWS):
        for c in range(COLS):

            cell = ws.cell(row=r+3, column=c+1)

            cell.value = seats[r][c]
            cell.alignment = align
            cell.border = border

            ws.column_dimensions[chr(65+c)].width = 15

        ws.row_dimensions[r+3].height = 25

    wb.save(save_path)

    messagebox.showinfo("成功", "Excel 已导出")


# ======================
# GUI
# ======================

root = tk.Tk()
root.title("智慧班级座位管理系统")
root.geometry("760x520")

top_frame = tk.Frame(root)
top_frame.pack(pady=10)

btn_load = tk.Button(
    top_frame,
    text="导入Excel名单",
    command=load_excel,
    width=15
)

btn_load.grid(row=0,column=0,padx=10)

btn_generate = tk.Button(
    top_frame,
    text="随机排座",
    command=generate_seats,
    width=15
)

btn_generate.grid(row=0,column=1,padx=10)

btn_export = tk.Button(
    top_frame,
    text="导出Excel",
    command=export_excel,
    width=15
)

btn_export.grid(row=0,column=2,padx=10)

status_label = tk.Label(root,text="请导入学生名单")
status_label.pack()

canvas = tk.Canvas(root,width=720,height=400)
canvas.pack()

root.mainloop()