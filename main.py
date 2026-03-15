import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import time
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from fpdf import FPDF
import os
import sys
import pystray
from PIL import Image, ImageDraw
import threading

# ======================
# 基础设置
# ======================
ROWS, COLS, TOTAL = 5, 6, 30
seat_names, seats = [], []
duty_tasks = ["擦黑板","教室扫拖","倒垃圾","室外卫生区","组长，负责各种擦(doge)"]
duty_students = []
duty_index = 0
history_file = "Duty_History.xlsx"
seat_file = "seat_layout.xlsx"
drag_data = {"widget":None,"row":0,"col":0}

# ======================
# 座位功能（同前版）
# ======================
def load_seat_excel():
    global seat_names
    file_path = filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")])
    if not file_path: return
    try:
        df = pd.read_excel(file_path, header=None)
        seat_names = df.iloc[:,0].dropna().tolist()
        messagebox.showinfo("成功",f"读取 {len(seat_names)} 名学生")
    except Exception as e:
        messagebox.showerror("错误",str(e))

def generate_seats(gender_mode=False, save_after=True):
    global seats
    if not seat_names:
        messagebox.showwarning("提示","请先导入座位名单")
        return
    if len(seat_names)>TOTAL:
        messagebox.showwarning("提示",f"人数超过 {TOTAL}")
        return
    shuffled = seat_names[:]
    random.shuffle(shuffled)
    if gender_mode:
        male = [s for s in shuffled if "(M)" in s]
        female = [s for s in shuffled if "(F)" in s]
        shuffled=[]
        while male or female:
            if male: shuffled.append(male.pop(0))
            if female: shuffled.append(female.pop(0))
        while len(shuffled)<TOTAL: shuffled.append("")
    else:
        while len(shuffled)<TOTAL: shuffled.append("")
    seats=[shuffled[i*COLS:(i+1)*COLS] for i in range(ROWS)]
    draw_seats()
    if save_after:
        save_seat_layout()

def save_seat_layout():
    if not seats: return
    wb=Workbook()
    ws=wb.active
    for r in range(ROWS):
        for c in range(COLS):
            ws.cell(r+1,c+1).value=seats[r][c]
    wb.save(seat_file)

def load_seat_layout():
    global seats
    if os.path.exists(seat_file):
        wb = load_workbook(seat_file)
        ws = wb.active
        seats = []
        for r in range(ROWS):
            row = []
            for c in range(COLS):
                val = ws.cell(r+1, c+1).value
                row.append(val if val is not None else "")
            seats.append(row)
        while len(seats) < ROWS:
            seats.append([""] * COLS)
        for r in range(ROWS):
            while len(seats[r]) < COLS:
                seats[r].append("")
    else:
        seats = [[""]*COLS for _ in range(ROWS)]

def on_click(event):
    row = int((event.y-100)//50)
    col = int((event.x-40)//110)
    if 0<=row<ROWS and 0<=col<COLS:
        drag_data["widget"]=seats[row][col]
        drag_data["row"], drag_data["col"]=row,col

def on_release(event):
    row = int((event.y-100)//50)
    col = int((event.x-40)//110)
    if drag_data["widget"] and 0<=row<ROWS and 0<=col<COLS:
        seats[drag_data["row"]][drag_data["col"]], seats[row][col]=seats[row][col], drag_data["widget"]
        draw_seats()
        save_seat_layout()
    drag_data["widget"]=None

def draw_seats():
    canvas.delete("all")
    cell_w, cell_h = 110, 50
    start_x, start_y = 40, 100
    canvas.create_rectangle(start_x, 30, start_x+COLS*cell_w, 70, fill="#dddddd")
    canvas.create_text(start_x+COLS*cell_w/2, 50, text="讲台", font=("Arial",14,"bold"))

    for r in range(ROWS):
        for c in range(COLS):
            x1, y1 = start_x + c*cell_w, start_y + r*cell_h
            x2, y2 = x1 + cell_w, y1 + cell_h
            canvas.create_rectangle(x1, y1, x2, y2)
            if r < len(seats) and c < len(seats[r]):
                name = seats[r][c] or ""
            else:
                name = ""
            canvas.create_text((x1+x2)/2, (y1+y2)/2, text=name, font=("Arial",10))

def export_seats_excel():
    if not seats:
        messagebox.showwarning("提示","请先生成座位表")
        return
    save_path=filedialog.asksaveasfilename(defaultextension=".xlsx")
    if not save_path: return
    wb=Workbook()
    ws=wb.active
    ws.title="座位表"
    border=Border(left=Side(style="thin"),right=Side(style="thin"),
                  top=Side(style="thin"),bottom=Side(style="thin"))
    align=Alignment(horizontal="center",vertical="center")
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=COLS)
    ws.cell(1,1).value="讲台"
    ws.cell(1,1).alignment=align
    ws.cell(1,1).font=Font(size=14,bold=True)
    for r in range(ROWS):
        for c in range(COLS):
            cell=ws.cell(row=r+3,column=c+1)
            cell.value=seats[r][c]
            cell.alignment=align
            cell.border=border
            ws.column_dimensions[chr(65+c)].width=15
        ws.row_dimensions[r+3].height=25
    wb.save(save_path)
    messagebox.showinfo("成功","座位表已导出")

# ======================
# 值日功能
# ======================
def load_duty_excel():
    global duty_students
    file_path=filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx *.xls")])
    if not file_path: return
    try:
        df=pd.read_excel(file_path,header=None)
        duty_students=df.iloc[:,0].dropna().tolist()
        messagebox.showinfo("成功",f"读取 {len(duty_students)} 名学生")
    except Exception as e:
        messagebox.showerror("错误",str(e))

def get_today_duty():
    global duty_index
    if not duty_students: return []
    selected=[]
    for i in range(5):
        selected.append(duty_students[duty_index % len(duty_students)])
        duty_index += 1
    return selected

def show_duty_window():
    today_duty=get_today_duty()
    if not today_duty:
        messagebox.showwarning("提示","请先导入值日名单")
        return
    duty_text=""
    roles=["甲","乙","丙","丁","戊"]
    for r,s,t in zip(roles,today_duty,duty_tasks):
        duty_text+=f"{r}：{s} {t}\n"
    duty_win=tk.Toplevel()
    duty_win.title("今日值日安排")
    duty_win.geometry("400x200")
    tk.Label(duty_win,text=duty_text,font=("Arial",12),justify="left").pack(pady=20)
    save_duty_history(today_duty)
    save_duty_pdf(today_duty)

def save_duty_history(today_duty):
    date_str=datetime.now().strftime("%Y-%m-%d")
    if os.path.exists(history_file):
        wb=load_workbook(history_file)
        ws=wb.active
    else:
        wb=Workbook()
        ws=wb.active
        ws.title="值日历史"
        ws.append(["日期","甲","乙","丙","丁","戊"])
    if ws.max_row>1 and ws.cell(ws.max_row,1).value==date_str: return
    ws.append([date_str]+today_duty)
    wb.save(history_file)

def save_duty_pdf(today_duty):
    pdf=FPDF()
    pdf.add_page()
    pdf.set_font("Arial","B",16)
    pdf.cell(0,10,f"值日安排 {datetime.now().strftime('%Y-%m-%d')}",ln=1,align="C")
    pdf.set_font("Arial","",14)
    roles=["甲","乙","丙","丁","戊"]
    for r,s,t in zip(roles,today_duty,duty_tasks):
        pdf.cell(0,10,f"{r}：{s} {t}",ln=1)
    pdf_file="Duty_"+datetime.now().strftime("%Y%m%d")+".pdf"
    pdf.output(pdf_file)

def duty_timer_thread():
    while True:
        now=time.localtime()
        if now.tm_hour==7 and now.tm_min==5:
            root.after(0,show_duty_window)
            time.sleep(60)
        time.sleep(20)

# ======================
# 系统托盘图标
# ======================
def create_image():
    # 生成一个简单图标
    image = Image.new('RGB', (64, 64), color=(255, 0, 0))
    d = ImageDraw.Draw(image)
    d.rectangle((16,16,48,48), fill=(255,255,255))
    return image

def on_quit(icon, item):
    icon.stop()
    root.quit()

def on_show(icon, item):
    root.deiconify()

def start_tray():
    image = create_image()
    menu = pystray.Menu(
        pystray.MenuItem("显示主窗口", on_show),
        pystray.MenuItem("退出", on_quit)
    )
    icon = pystray.Icon("SmartClassroom", image, "智慧班级", menu)
    icon.run()

# ======================
# GUI 初始化
# ======================
root = tk.Tk()
root.title("智慧班级管理系统 终极升级版")
root.geometry("820x550")

# 判断是否开机启动
if len(sys.argv) > 1 and sys.argv[1] == "--startup":
    root.withdraw()  # 开机启动隐藏
    threading.Thread(target=start_tray, daemon=True).start()
else:
    root.deiconify()  # 正常启动显示

top_frame=tk.Frame(root)
top_frame.pack(pady=10)

tk.Button(top_frame,text="导入座位名单",command=load_seat_excel,width=15).grid(row=0,column=0,padx=5)
tk.Button(top_frame,text="随机排座",command=lambda:generate_seats(gender_mode=False),width=15).grid(row=0,column=1,padx=5)
tk.Button(top_frame,text="男女交错排座",command=lambda:generate_seats(gender_mode=True),width=15).grid(row=0,column=2,padx=5)
tk.Button(top_frame,text="导出座位表",command=export_seats_excel,width=15).grid(row=0,column=3,padx=5)
tk.Button(top_frame,text="导入值日名单",command=load_duty_excel,width=15).grid(row=0,column=4,padx=5)

tk.Label(root,text="值日安排已启用，每天7:05自动弹窗并记录历史/PDF，拖拽座位会自动保存",font=("Arial",10)).pack()
canvas=tk.Canvas(root,width=780,height=400)
canvas.pack()
canvas.bind("<Button-1>",on_click)
canvas.bind("<ButtonRelease-1>",on_release)

# 自动加载座位布局
load_seat_layout()
draw_seats()

# 值日线程
threading.Thread(target=duty_timer_thread, daemon=True).start()

root.mainloop()
