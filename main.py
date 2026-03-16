import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random
import threading
import time
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side, Font
from fpdf import FPDF
import os
import sys
import pystray
from PIL import Image, ImageDraw

ROWS, COLS, TOTAL = 5, 6, 30
seat_names = []
seats = []

duty_tasks = ["擦黑板","教室扫拖","倒垃圾","室外卫生区","组长，负责各种擦(doge)"]
duty_students = []
duty_queue = []

seat_file = "seat_layout.xlsx"
history_file = "Duty_History.xlsx"

drag_data={"name":None,"row":0,"col":0}

# =========================
# 座位功能
# =========================

def load_seat_excel():
    global seat_names
    file=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
    if not file:return
    df=pd.read_excel(file,header=None)
    seat_names=df.iloc[:,0].dropna().tolist()
    messagebox.showinfo("成功",f"读取 {len(seat_names)} 名学生")

def generate_seats():
    global seats
    if not seat_names:
        messagebox.showwarning("提示","请先导入名单")
        return
    shuffled=seat_names[:]
    random.shuffle(shuffled)
    while len(shuffled)<TOTAL:
        shuffled.append("")
    seats=[shuffled[i*COLS:(i+1)*COLS] for i in range(ROWS)]
    draw_seats()
    save_seat_layout()

def draw_seats():
    canvas.delete("all")

    cell_w=110
    cell_h=50
    start_x=40
    start_y=100

    canvas.create_rectangle(start_x,30,start_x+COLS*cell_w,70,fill="#ddd")
    canvas.create_text(start_x+COLS*cell_w/2,50,text="讲台",font=("Arial",14,"bold"))

    for r in range(ROWS):
        for c in range(COLS):

            x1=start_x+c*cell_w
            y1=start_y+r*cell_h
            x2=x1+cell_w
            y2=y1+cell_h

            canvas.create_rectangle(x1,y1,x2,y2)

            name=""
            if r<len(seats) and c<len(seats[r]):
                name=seats[r][c] or ""

            canvas.create_text((x1+x2)/2,(y1+y2)/2,text=name,font=("Arial",10))

def save_seat_layout():
    wb=Workbook()
    ws=wb.active

    for r in range(ROWS):
        for c in range(COLS):
            ws.cell(r+1,c+1).value=seats[r][c]

    wb.save(seat_file)

# =========================
# 值日功能（随机轮换）
# =========================

def load_duty_excel():
    global duty_students
    file=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
    if not file:return
    df=pd.read_excel(file,header=None)
    duty_students=df.iloc[:,0].dropna().tolist()
    messagebox.showinfo("成功",f"读取 {len(duty_students)} 名学生")

def get_today_duty():

    global duty_queue

    if not duty_students:
        return []

    selected=[]

    while len(selected)<5:

        if len(duty_queue)==0:
            duty_queue=duty_students[:]
            random.shuffle(duty_queue)

        selected.append(duty_queue.pop(0))

    return selected

def show_duty_window():

    today=get_today_duty()

    if not today:
        messagebox.showwarning("提示","请先导入值日名单")
        return

    roles=["甲","乙","丙","丁","戊"]

    text=""

    for r,s,t in zip(roles,today,duty_tasks):
        text+=f"{r}：{s} {t}\n"

    win=tk.Toplevel()
    win.title("今日值日")
    win.geometry("400x200")

    tk.Label(win,text=text,font=("Arial",12),justify="left").pack(pady=20)

    save_history(today)
    save_pdf(today)

def save_history(today):

    date=datetime.now().strftime("%Y-%m-%d")

    if os.path.exists(history_file):
        wb=load_workbook(history_file)
        ws=wb.active
    else:
        wb=Workbook()
        ws=wb.active
        ws.append(["日期","甲","乙","丙","丁","戊"])

    if ws.max_row>1 and ws.cell(ws.max_row,1).value==date:
        return

    ws.append([date]+today)

    wb.save(history_file)

def save_pdf(today):

    pdf=FPDF()
    pdf.add_page()

    pdf.set_font("Arial","B",16)
    pdf.cell(0,10,"值日安排",ln=1,align="C")

    pdf.set_font("Arial","",14)

    roles=["甲","乙","丙","丁","戊"]

    for r,s,t in zip(roles,today,duty_tasks):
        pdf.cell(0,10,f"{r}：{s} {t}",ln=1)

    pdf.output("Duty_"+datetime.now().strftime("%Y%m%d")+".pdf")

# =========================
# 自动弹窗线程
# =========================

def duty_timer_thread():

    while True:

        now=time.localtime()

        if now.tm_hour==20 and now.tm_min==40:

            root.after(0,show_duty_window)

            time.sleep(60)

        time.sleep(20)

# =========================
# 托盘图标
# =========================

def create_icon():

    image=Image.new("RGB",(64,64),(255,0,0))
    d=ImageDraw.Draw(image)
    d.rectangle((16,16,48,48),fill=(255,255,255))

    return image

def tray_show(icon,item):
    root.deiconify()

def tray_quit(icon,item):
    icon.stop()
    root.quit()

def start_tray():

    icon=pystray.Icon(
        "SmartClass",
        create_icon(),
        "智慧班级",
        menu=pystray.Menu(
            pystray.MenuItem("显示主窗口",tray_show),
            pystray.MenuItem("退出",tray_quit)
        )
    )

    icon.run()

# =========================
# GUI
# =========================

root=tk.Tk()
root.title("智慧班级管理系统")
root.geometry("820x550")

if len(sys.argv)>1 and sys.argv[1]=="--startup":
    root.withdraw()
    threading.Thread(target=start_tray,daemon=True).start()

top_frame=tk.Frame(root)
top_frame.pack(pady=10)

tk.Button(top_frame,text="导入座位名单",command=load_seat_excel,width=15).grid(row=0,column=0,padx=5)
tk.Button(top_frame,text="随机排座",command=generate_seats,width=15).grid(row=0,column=1,padx=5)
tk.Button(top_frame,text="导入值日名单",command=load_duty_excel,width=15).grid(row=0,column=2,padx=5)
tk.Button(top_frame,text="生成今日值日",command=show_duty_window,width=15).grid(row=0,column=3,padx=5)

canvas=tk.Canvas(root,width=780,height=400)
canvas.pack()

draw_seats()

threading.Thread(target=duty_timer_thread,daemon=True).start()

root.mainloop()
