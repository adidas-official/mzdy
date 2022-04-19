import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
from months_cz import months_cz

current_month = months_cz[datetime.now().month - 1]


def main_window(widget, width=0, height=0):

    screen_w, screen_h = widget.winfo_screenwidth(), widget.winfo_screenheight()

    left = int(screen_w / 2) - int(width / 2)
    top = int(screen_h / 2) - int(height / 2)

    if width and height:
        widget.geometry(f'{width}x{height}+{left}+{top}')
    else:
        widget.geometry('+%d+%d' %(500, 100))


root = tk.Tk()
root.grid_columnconfigure(0, weight=1)

main_window(root, 720, 510)
s = ttk.Style()
s.configure("GrooveBorder.TFrame", relief='groove')
opts = {'padx': 10, 'sticky': 'WE', 'ipadx': 10, 'ipady': 10}

top_frame = ttk.Frame(root, height=180, style="GrooveBorder.TFrame")

top_frame.grid_columnconfigure(0, weight=1)
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_columnconfigure(2, weight=2)

but1 = ttk.Button(top_frame, width=1)
but2 = ttk.Button(top_frame, width=1)
but3 = ttk.Button(top_frame, width=1)
txt = ScrolledText(top_frame, width=1, height=1)

btn_opts = {'sticky': 'we', 'padx': 5, 'pady': 5}

but1.grid(row=0, column=0, **btn_opts)
but2.grid(row=1, column=0, **btn_opts)
but3.grid(row=0, column=1, **btn_opts)
txt.grid(row=0, column=2, rowspan=2, **btn_opts)

top_frame.grid(row=0, column=0, **opts, pady=10)

bottom_frame = ttk.Frame(root, height=300, style='GrooveBorder.TFrame')
bottom_frame.grid(row=1, column=0, **opts)

root.mainloop()