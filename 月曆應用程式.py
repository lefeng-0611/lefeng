import ctypes
import sys
import tkinter as tk
import calendar
from datetime import datetime
from tkinter import messagebox
from PIL import Image, ImageTk
import requests
from io import BytesIO
import json
import os

# ---------------- 自動提權 ----------------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{sys.argv[0]}"', None, 1)
    sys.exit()

# ---------------- 行事曆應用 ----------------
class CalendarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("月曆應用程式")

        # 視窗圖標設定
        try:
            url = "https://img95.699pic.com/xsj/08/xg/pp.jpg!/fw/700/watermark/url/L3hzai93YXRlcl9kZXRhaWwyLnBuZw/align/southeast"
            response = requests.get(url)
            img_data = response.content
            img = Image.open(BytesIO(img_data)).resize((32, 32))
            self.root.iconphoto(False, ImageTk.PhotoImage(img))
        except:
            pass

        self.root.geometry("520x620")
        self.root.config(bg="#2196F3")

        current_date = datetime.today()
        self.year = current_date.year
        self.month = current_date.month
        self.today = current_date.day

        self.event_file = "events.json"
        self.events = self.load_events()

        self.calendar_frame = tk.Frame(self.root, bg="#2196F3")
        self.calendar_frame.pack(pady=20)

        self.month_label = tk.Label(self.calendar_frame, font=("Arial", 14), bg="#2196F3", fg="white")
        self.month_label.grid(row=0, column=0, padx=10)

        self.year_label = tk.Label(self.calendar_frame, font=("Arial", 14), bg="#2196F3", fg="white")
        self.year_label.grid(row=0, column=1, padx=10)

        self.display_calendar()

        self.controls_frame = tk.Frame(self.root, bg="#2196F3")
        self.controls_frame.pack()

        self.prev_month_button = tk.Button(self.controls_frame, text="上一月", font=("Arial", 12), command=self.prev_month, bg="#D32F2F", fg="white", relief="solid", width=8)
        self.prev_month_button.grid(row=0, column=0, padx=10, pady=10)

        self.next_month_button = tk.Button(self.controls_frame, text="下一月", font=("Arial", 12), command=self.next_month, bg="#D32F2F", fg="white", relief="solid", width=8)
        self.next_month_button.grid(row=0, column=1, padx=10, pady=10)

        self.prev_year_button = tk.Button(self.controls_frame, text="上一年", font=("Arial", 12), command=self.prev_year, bg="#1976D2", fg="white", relief="solid", width=8)
        self.prev_year_button.grid(row=1, column=0, padx=10, pady=10)

        self.next_year_button = tk.Button(self.controls_frame, text="下一年", font=("Arial", 12), command=self.next_year, bg="#1976D2", fg="white", relief="solid", width=8)
        self.next_year_button.grid(row=1, column=1, padx=10, pady=10)

    def display_calendar(self):
        for widget in self.calendar_frame.winfo_children():
            widget.destroy()

        month_names = ["", "一月", "二月", "三月", "四月", "五月", "六月",
                       "七月", "八月", "九月", "十月", "十一月", "十二月"]
        year_month = f"{self.year} 年 {month_names[self.month]}"

        label = tk.Label(self.calendar_frame, text=year_month, font=("Arial", 16, "bold"), bg="#2196F3", fg="Black")
        label.grid(row=1, column=0, columnspan=7, pady=10)

        month_calendar = calendar.monthcalendar(self.year, self.month)
        week_days = ["一", "二", "三", "四", "五", "六", "日"]
        for i, day in enumerate(week_days):
            day_label = tk.Label(self.calendar_frame, text=day, font=("Arial", 12, "bold"), bg="#2196F3", fg="#1C1C1C")
            day_label.grid(row=2, column=i, padx=5, pady=5)

        for row, week in enumerate(month_calendar, start=3):
            for col, day in enumerate(week):
                if day != 0:
                    color = "white"
                    if day == self.today and self.year == datetime.today().year and self.month == datetime.today().month:
                        color = "yellow"
                    elif col in (5, 6):
                        color = "#03A9F4"

                    has_event = f"{self.year}-{self.month:02d}-{day:02d}" in self.events
                    btn = tk.Button(
                        self.calendar_frame, text=str(day), width=4, height=2,
                        relief="solid", font=("Arial", 12),
                        bg=color if not has_event else "#EAFF00",
                        fg="black" if not has_event else "#E91E63",
                        command=lambda d=day: self.on_date_selected(d)
                    )
                    btn.grid(row=row, column=col, padx=5, pady=5)

    def on_date_selected(self, day):
        selected_date = f"{self.year}-{self.month:02d}-{day:02d}"
        top = tk.Toplevel(self.root)
        top.title("事件記事")
        top.geometry("300x200")

        tk.Label(top, text=selected_date, font=("Arial", 12)).pack(pady=10)
        event_text = tk.Text(top, height=5, width=30)
        event_text.pack()
        event_text.insert("1.0", self.events.get(selected_date, ""))

        def save_event():
            text = event_text.get("1.0", tk.END).strip()
            if text:
                self.events[selected_date] = text
            elif selected_date in self.events:
                del self.events[selected_date]
            self.save_events()
            self.display_calendar()
            top.destroy()

        tk.Button(top, text="儲存", command=save_event).pack(pady=5)

    def save_events(self):
        with open(self.event_file, 'w', encoding='utf-8') as f:
            json.dump(self.events, f, ensure_ascii=False, indent=2)

    def load_events(self):
        if os.path.exists(self.event_file):
            with open(self.event_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        return {}

    def prev_month(self):
        if self.month == 1:
            self.month = 12
            self.year -= 1
        else:
            self.month -= 1
        self.display_calendar()

    def next_month(self):
        if self.month == 12:
            self.month = 1
            self.year += 1
        else:
            self.month += 1
        self.display_calendar()

    def prev_year(self):
        self.year -= 1
        self.display_calendar()

    def next_year(self):
        self.year += 1
        self.display_calendar()

if __name__ == "__main__":
    root = tk.Tk()
    app = CalendarApp(root)
    root.mainloop()