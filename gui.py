import time
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
from months_cz import months_cz
from tkinter.filedialog import askopenfilename
from pathlib import Path
from time import sleep
import json


with open('structure.json', 'r', encoding='cp1250') as jdata:
    last_data = json.load(jdata)

current_month = months_cz[datetime.now().month - 1]


def main_window(widget, width=0, height=0):

    screen_w, screen_h = widget.winfo_screenwidth(), widget.winfo_screenheight()

    left = int(screen_w / 2) - int(width / 2)
    top = int(screen_h / 2) - int(height / 2)

    if width and height:
        widget.geometry(f'{width}x{height}+{left}+{top}')
    else:
        widget.geometry('+%d+%d' % (500, 100))


root = tk.Tk()
root.grid_columnconfigure(0, weight=1)

main_window(root, 720, 510)
opts = {'padx': 10, 'sticky': 'WE', 'ipadx': 10, 'ipady': 10}

top_frame = ttk.Frame(root, height=180, style="GrooveBorder.TFrame")

top_frame.grid_columnconfigure(0, weight=1)
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_columnconfigure(2, weight=2)


def start():
    for x in range(50):
        progress['value'] += 2
        root.update_idletasks()
        sleep(0.1)
    progress.stop()



def open_log_file():
    progress.stop()


def select_log_file():
    filetypes = [
        ('log file', '*.log')
    ]

    filename = askopenfilename(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes
    )

    if filename:
        file = Path(filename).name
        # files['log file'] = filename
        print(file)
        log_file_btn['text'] = file
    else:
        log_file_btn['text'] = 'Choose log file'


def set_datas(comp, filetypes, key_name):
    # print(comp)  # .!frame2.!notebook.!frame, .!frame2.!notebook.!frame2
    c_name = tabs.tab(comp)["text"]
    # print(company_name)  # Fiala, Bereko

    filename = askopenfilename(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes
    )

    if filename:
        # btn = f'{comp}.!button5'
        file = Path(filename).name
        # root.nametowidget(btn)["text"] = file
        last_data[c_name][key_name] = str(Path(filename))
        with open('structure.json', 'w', encoding='cp1250') as input_file:
            input_file.write(json.dumps(last_data))


chosen_month = tk.StringVar()
choose_month = ttk.OptionMenu(top_frame, chosen_month, current_month, *months_cz)
choose_month.configure(width=8)

progress = ttk.Progressbar(top_frame, length=100, mode='determinate', orient='horizontal')

log_file_btn = ttk.Button(top_frame, width=1, text='Open log file', command=progress.stop)

start_btn = ttk.Button(top_frame, width=1, text='Start', state='enabled', command=start)

txt = ScrolledText(top_frame, width=1, height=5)

btn_opts = {'sticky': 'we', 'padx': 5, 'pady': 5}

choose_month.grid(row=0, column=0, **btn_opts)
progress.grid(row=1, column=1, **btn_opts)
start_btn.grid(row=1, column=0, **btn_opts)
log_file_btn.grid(row=0, column=1, **btn_opts)
txt.grid(row=0, column=2, rowspan=2, **btn_opts)

top_frame.grid(row=0, column=0, **opts, pady=10)

bottom_frame = ttk.Frame(root, height=300)
bottom_frame.grid(row=1, column=0, **opts)

tabs = ttk.Notebook(bottom_frame, width=700, height=300)
tabs.grid(row=0, column=0)

# companies = last_data.keys()
for company_name, file_paths in last_data.items():
    # print(company_name)
    # print(file_paths)

    # files.setdefault(company_name, {})

    company_frame = ttk.Frame(tabs)
    company_frame.grid_columnconfigure(0, weight=1)
    company_frame.grid_columnconfigure(1, weight=1)
    company_frame.grid_columnconfigure(2, weight=1)
    company_frame.grid_columnconfigure(3, weight=1)

    check_box = ttk.Checkbutton(company_frame, text=company_name, onvalue=1, offvalue=0)
    check_box.grid(row=0, column=0)

    src_label = ttk.Label(company_frame, text='ZDROJ')
    src_label.grid(row=1, column=0)

    up_label_in = ttk.Label(company_frame, text='pro u.p.')
    up_label_in.grid(row=2, column=0)

    up_btn_in = ttk.Button(
        company_frame,
        text=Path(file_paths['src_file_up']).name,
        width=20,
        command=lambda: set_datas(
            tabs.select(),
            [('Spreadsheets', '*.xlsx')],
            'src_file_up')
    )

    up_btn_in.grid(row=3, column=0)

    inter_label_in = ttk.Label(company_frame, text='interni')
    inter_label_in.grid(row=4, column=0)

    inter_btn_in = ttk.Button(
        company_frame,
        text=Path(file_paths['src_file_loc']).name,
        width=20,
        command=lambda: set_datas(
            tabs.select(),
            [('Spreadsheets', '*.xlsx')],
            'src_file_loc')
    )

    inter_btn_in.grid(row=5, column=0)

    output_label = ttk.Label(company_frame, text='VYSTUP')
    output_label.grid(row=1, column=1)

    up_label_out = ttk.Label(company_frame, text='pro u.p.')
    up_label_out.grid(row=2, column=1)

    up_btn_out = ttk.Button(
        company_frame,
        text=file_paths['output_file_up'],
        width=20,
        command=lambda: set_datas(
            tabs.select(),
            [('Excel sheet', '*.xlsx'),
             ('Excel 2010', '*.xls'),
             ('Open document', '*.ods')],
            'output_file_up'
        )
    )

    up_btn_out.grid(row=3, column=1)

    inter_label_out = ttk.Label(company_frame, text='interni')
    inter_label_out.grid(row=4, column=1)

    inter_btn_out = ttk.Button(
        company_frame,
        text=file_paths['output_file_loc'],
        width=20,
        command=lambda: set_datas(
            tabs.select(),
            [('Spreadsheets', '*.xlsx')],
            'output_file_loc'
        )
    )

    inter_btn_out.grid(row=5, column=1)

    data_label = ttk.Label(company_frame, text='Data')
    data_label.grid(row=0, column=2)

    # data_btn = ttk.Button(company_frame, text='input data', width=20, command=lambda: set_data(tabs.tab(tabs.select())))
    data_btn = ttk.Button(
        company_frame,
        text=Path(file_paths['input_data']).name,
        width=20,
        command=lambda: set_datas(
            tabs.select(),
            [('CSV', '*.csv')],
            'input_data'
        )
    )

    data_btn.grid(row=0, column=3)

    ins_groups_label = ttk.Label(company_frame, text='POJISTOVNY')
    ins_groups_label.grid(row=1, column=2)

    tree_label = ttk.Label(company_frame, text='kody zdravotnich pojistoven v ucto2000')
    tree_label.grid(row=2, column=2)
    tree_cols = ('ucto_code', 'ins_code', 'ins_name')
    tree = ttk.Treeview(company_frame, columns=tree_cols, show='headings')
    tree.column('ucto_code', width=70)
    tree.column('ins_code', width=100)
    tree.heading('ucto_code', text='Ucto2000')
    tree.heading('ins_code', text='kod pojistovny')
    tree.heading('ins_name', text='nazev pojistovny')
    tree.grid(row=3, column=2, columnspan=2, rowspan=8)
    for widget in company_frame.winfo_children():
        widget.grid(padx=10, pady=10, sticky='NW')

    tabs.add(company_frame, text=company_name)
tabs.add(ttk.Frame(tabs), text='+')

root.mainloop()
