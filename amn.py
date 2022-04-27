# amn.py
# -------------+
# Automatizace |
# Mzdovych     |
# Nakladu      |
# -------------+
# v2
# author: Zdenek Frydryn
# created for: Bereko s.r.o., Drogerie Fiala s.r.o.,
# Description: Automatizace vyplnovani mezd do xlsx tabulek pro urad prace a internich tabulek.

# TOOD:
# 1) Check for quartal => [DONE]
# 2) mzdove naklady => [DONE]
# 3) choose bereko or fiala from filename of csv file; bereko.csv or fiala.csv => [DONE]
# 4) update both UP and mzdove naklady in one run => [DONE]
# 5) update both fiala and bereko in one run => [DONE]
# 6) GUI for updating structure.json and for executing program => [DONE]
# 7) insurance codes => [DONE]
# 8) copy status from month before
# 9) if first month of quartal, delete next 2 months
# 10) add `Davky1` field to csv export => [DONE]
# --------------------------------------------------------------------------------
# maybe not possible:
# get status of ozp, add to exported csv


import logging
import csv
import openpyxl
import msoffcrypto
from pathlib import Path
from openpyxl.utils import column_index_from_string
from months_cz import months_cz
from io import BytesIO
import json
# from shutil import copyfile
import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from datetime import datetime
from tkinter.filedialog import askopenfilename

logging.basicConfig(level=logging.INFO, filename='log.log', filemode='w',
                    format='%(levelname)s - %(asctime)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')
current_month = months_cz[datetime.now().month - 1]


def open_log_file():
    pass


def load_ins_codes(c_name, my_tree):
    my_tree.delete(*my_tree.get_children())
    ins_codes_file = f'insurance_codes_{c_name}.json'
    if Path(ins_codes_file).exists():
        ins_codes = json.load(open(ins_codes_file))
        for ins_code, data in ins_codes.items():
            my_tree.insert('', tk.END, values=[ins_code, data[0], data[1]])


def add_new(tab):
    txt.delete('1.0', tk.END)
    my_tree = root.nametowidget(tab + '.!treeview')
    c_name = tabs.tab(tab)['text']

    ins_codes_file = f'insurance_codes_{c_name}.json'
    ins_codes = []
    i_codes = []

    if Path(ins_codes_file).exists():
        ins_codes = json.load(open(ins_codes_file))
        for i_code in ins_codes.values():
            i_codes.append(i_code[0])

    ucto_code_entry = str(root.nametowidget(tab + '.!entry').get())
    ins_code_entry = str(root.nametowidget(tab + '.!entry2').get())
    ins_name_entry = root.nametowidget(tab + '.!entry3').get()

    values = (ucto_code_entry, ins_code_entry, ins_name_entry)

    # Conditions
    if ucto_code_entry and ins_code_entry and ins_name_entry:
        if ucto_code_entry not in ins_codes and ins_code_entry not in i_codes:
            my_tree.insert('', tk.END, values=values)
        else:
            txt.insert('1.0', 'Tento kod uz existuje.')
            print('This code is already used')
    else:
        txt.insert('1.0', 'Vyplnte vsechna pole.')
        print('Fill out all fields')

    save_ins(tab, my_tree)


def update_ins(tab, my_tree):
    selected = my_tree.focus()
    ucto_code_entry = str(root.nametowidget(tab + '.!entry').get())
    ins_code_entry = str(root.nametowidget(tab + '.!entry2').get())
    ins_name_entry = root.nametowidget(tab + '.!entry3').get()

    values = (ucto_code_entry, ins_code_entry, ins_name_entry)
    my_tree.item(selected, text="", values=values)

    save_ins(tab, my_tree)


def delete_record(tab, my_tree):
    selected = my_tree.focus()
    my_tree.delete(selected)
    save_ins(tab, my_tree)


def save_ins(tab, my_tree):
    c_name = tabs.tab(tab)['text']
    ins_data = {}
    for row in my_tree.get_children():
        item = my_tree.item(row)['values']
        ucto_code = str(item[0]).zfill(5)
        ins_group_code = str(item[1])
        ins_name = item[2]
        ins_data.setdefault(ucto_code, [])
        ins_data[ucto_code].append(ins_group_code)
        ins_data[ucto_code].append(ins_name)

    ins_codes_file = f'insurance_codes_{c_name}.json'
    if not Path(ins_codes_file).exists():
        Path(ins_codes_file).touch()

    with open(f'insurance_codes_{c_name}.json', 'w') as outfile:
        json.dump(ins_data, outfile)


# noinspection PyUnusedLocal
def item_selected(event, my_tree):
    c_frame = my_tree.winfo_parent()
    ucto_code_entry = root.nametowidget(c_frame + '.!entry')
    ins_code_entry = root.nametowidget(c_frame + '.!entry2')
    ins_name_entry = root.nametowidget(c_frame + '.!entry3')

    ucto_code_entry.delete(0, tk.END)
    ins_code_entry.delete(0, tk.END)
    ins_name_entry.delete(0, tk.END)

    selected_item = my_tree.focus()
    row = my_tree.item(selected_item)
    record = row['values']

    ucto_code_entry.insert(0, str(record[0]).zfill(5))
    ins_code_entry.insert(0, record[1])
    ins_name_entry.insert(0, record[2])


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


def set_datas(btn, comp, filetypes, key_name):
    # print(comp)  # .!frame2.!notebook.!frame, .!frame2.!notebook.!frame2
    c_name = tabs.tab(comp)["text"]
    # print(company_name)  # Fiala, Bereko

    filename = askopenfilename(
        title='Open a file',
        initialdir='.',
        filetypes=filetypes
    )

    if filename:
        file = Path(filename).name
        root.nametowidget(comp + '.' + btn.winfo_name())['text'] = file

        last_data[c_name][key_name] = str(Path(filename))
        with open('structure.json', 'w', encoding='cp1250') as input_file:
            input_file.write(json.dumps(last_data))


def amn(month_name, data, text_field):
    text_field.delete('1.0', tk.END)
    month = months_cz.index(month_name) + 1

    for c_name, input_output in data.items():
        try:
            with open(f'insurance_codes_{c_name}.json', 'r') as ins_file:
                ins_codes = json.load(ins_file)
        except Exception as e:
            print(e)

        text_field.insert(tk.END, f'|- Vyplnuji {c_name}\n')
        text_field.insert(tk.END, '=' * 44 + '\n')
        logging.info(f'Vyplnuji {c_name}')

        wb = openpyxl.load_workbook(input_output['src_file_up'])
        ws = wb["1) Úvodní list"]

        if month <= 3:
            ws.cell(row=6, column=4).value = 1
        elif 4 < month <= 6:
            ws.cell(row=6, column=4).value = 2
        elif 7 < month <= 9:
            ws.cell(row=6, column=4).value = 3
        else:
            ws.cell(row=6, column=4).value = 4

        with open(input_output['input_data'], 'r') as empl_data_csv:
            empl_data_reader = csv.reader(empl_data_csv)
            empl_data = list(empl_data_reader)
            empl_data = empl_data[1:]

        employees_up = {}
        employees_inter = {}

        # get data for each employee
        # split data to first name, last name, id, ins.group and money
        for i in empl_data:

            # get first name and last name
            name = i[0][1:-1]
            full_name = name.split(' ')
            lname = full_name[0]

            # Name might include title, merge first name with title
            if len(full_name) > 2:
                fname = " ".join(full_name[1:])
            else:
                fname = full_name[1]

            # get id number of employee
            id_num = i[1][1:-1].replace('/', '')

            # get code of insurance group
            ins = i[2][1:-1]
            if ins in ins_codes:
                ins_group_code = ins_codes[ins][0]
            else:
                ins_group_code = 999

            # get category of employment contract
            category = i[3][1:-1]

            # calculate salary
            total_expenses = 0
            for exp in i[5:]:
                try:
                    total_expenses += int(exp)
                except ValueError:
                    continue

            # noinspection PyTypeChecker
            employees_up.setdefault(id_num, {'first name': fname,
                                             'last name': lname,
                                             'ins code': ins_group_code,
                                             'cat': category,
                                             'payment expenses': total_expenses
                                             })

            if i[4]:
                expenses = (total_expenses, int(i[4]))
            else:
                expenses = (total_expenses, '')
            employees_inter.setdefault(lname + ' ' + fname, expenses)
            # print(employees_inter)

        # open mzdy UP table, go through each name in data and check if it is in table
        sheets = ['2) jmenný seznam', '3) nákl. prov. z. a prac. a.']
        col_letter_id = column_index_from_string('D')

        if month % 3 == 1:
            col_letter_pay = {sheets[0]: 11, sheets[1]: 9}
        elif month % 3 == 2:
            col_letter_pay = {sheets[0]: 16, sheets[1]: 10}
        else:
            col_letter_pay = {sheets[0]: 21, sheets[1]: 11}

        ids_in_xlsx = {}
        last_rows = []

        for sheet in sheets:
            progress['value'] = 0
            root.update_idletasks()
            ws = wb[sheet]
            last_row = 13
            ids_in_xlsx.setdefault(sheet, {})

            for i in range(100):
                row = i + 13
                id_of_person = ws.cell(row=row, column=col_letter_id).value
                # make list of ids in xlsx file
                if id_of_person:
                    ids_in_xlsx[sheet].setdefault(str(id_of_person).replace('/', ''), row)
                    last_row = row

            last_rows.append(last_row)

        for emp_id, emp_data in employees_up.items():
            progress['value'] += 100 / (len(employees_up))
            found = False
            for sheet_name, sheet_data in ids_in_xlsx.items():
                if emp_id in sheet_data:
                    found = True
            if not found:
                logging.info(f"Novy zamestnanec {emp_data['first name']} {emp_data['last name']}")
                text_field.insert(tk.END, f"|- Novy zam. {emp_data['first name']} {emp_data['last name']} ")
                text_field.see(tk.END)
                root.update_idletasks()
                if emp_data['cat'] == 'INV':
                    text_field.insert(tk.END, f"-> {sheets[0][:3]}\n")
                    logging.info(f"Pridavam do listu {sheets[0]}")
                    ws = wb[sheets[0]]
                    ws.cell(row=last_rows[0] + 1, column=2).value = emp_data['last name']
                    ws.cell(row=last_rows[0] + 1, column=3).value = emp_data['first name']
                    ws.cell(row=last_rows[0] + 1, column=4).value = emp_id
                    ws.cell(row=last_rows[0] + 1, column=7).value = emp_data['ins code']
                    ws.cell(row=last_rows[0] + 1, column=col_letter_pay[ws.title]).value = emp_data['payment expenses']
                    last_rows[0] += 1
                elif emp_data['cat'] == 'U':
                    text_field.insert(tk.END, ":Ucen\n")
                    logging.info("Ucen")
                    continue

                else:
                    text_field.insert(tk.END, f"-> {sheets[1][:3]}\n")
                    logging.info(f"Pridavam do listu {sheets[1]}")
                    ws = wb[sheets[1]]
                    ws.cell(row=last_rows[1] + 1, column=2).value = emp_data['last name']
                    ws.cell(row=last_rows[1] + 1, column=3).value = emp_data['first name']
                    ws.cell(row=last_rows[1] + 1, column=4).value = emp_id[:6] + '/' + emp_id[6:]
                    ws.cell(row=last_rows[1] + 1, column=col_letter_pay[ws.title]).value = emp_data['payment expenses']
                    ws.cell(row=last_rows[1] + 1, column=5).value = '-\'\'-'
                    ws.cell(row=last_rows[1] + 1, column=6).value = 'PA'
                    ws.cell(row=last_rows[1] + 1, column=7).value = '100%'
                    last_rows[1] += 1

        text_field.insert(tk.END, "=" * 44 + '\n')
        logging.info("-" * 44)

        for sheet_name, data in ids_in_xlsx.items():
            progress['value'] = 0
            ws = wb[sheet_name]

            for id_num, row_num in data.items():
                progress['value'] += 100 / (len(data.items()))
                if id_num in employees_up:
                    text_field.insert(tk.END,
                                      f"|- {employees_up[id_num]['first name']} {employees_up[id_num]['last name']}:{employees_up[id_num]['payment expenses']}\n")
                    logging.info(
                        f"Vyplnuji vyplatu pro {employees_up[id_num]['first name']} {employees_up[id_num]['last name']} v listu {sheet_name}:{employees_up[id_num]['payment expenses']}")
                    # zadat plat za tento mesic
                    ws.cell(row=row_num, column=col_letter_pay[sheet_name]).value = employees_up[id_num]['payment expenses']

                else:
                    text_field.insert(tk.END, f"|- Zamestnanec {id_num} v listu {sheet_name} nema vyplnenou vyplatu. Mazu pole status\n")
                    logging.info(f"Zamestnanec {id_num} v listu {sheet_name} nema vyplnenou vyplatu. Mazu pole status")
                    # smazat v xlsx status pro tento mesic
                    if sheet_name == sheets[0]:
                        ws.cell(row=row_num, column=col_letter_pay[sheet_name] - 1).value = ''
                root.update_idletasks()

        wb.save(input_output['output_file_up'])
        text_field.insert(tk.END, f"{c_name} vyplneno.\n")
        text_field.insert(tk.END, "=" * 44+"\n")
        logging.info(f"{c_name} vyplneno.\n")

        if 'src_file_loc' in input_output.keys():
            if Path(input_output['src_file_loc']).exists():
                try:
                    decrypted_wb = BytesIO()
                    with open(input_output['src_file_loc'], 'rb') as f:
                        office_file = msoffcrypto.OfficeFile(f)
                        office_file.load_key(password='13881744')
                        office_file.decrypt(decrypted_wb)

                    wb = openpyxl.load_workbook(filename=decrypted_wb)
                except UnboundLocalError:
                    wb = openpyxl.load_workbook(input_output['src_file_loc'])

                text_field.insert(tk.END, 'Vyplnuji data do interni tabulky\n')
                logging.info('Vyplnuji data do interni tabulky')
                fare_offset = 4

                for sheet in wb.sheetnames[:8]:
                    progress['value'] = 0
                    root.update_idletasks()
                    if sheet == 'Úřad práce':
                        fare_offset = 2
                    text_field.insert(tk.END, '\n+ ' + sheet + '\n')
                    logging.info(sheet)
                    ws = wb[sheet]
                    month_col = ''
                    for col in range(2, column_index_from_string('CH')):
                        cell = ws.cell(row=1, column=col)

                        if not type(cell).__name__ == 'MergedCell' and cell.value:
                            # if not type(cell).__name__ == 'MergedCell' and cell.value not in [None, "celkem za rok", "=A1"]:
                            if cell.value == month_name:
                                month_col = cell.column
                                break
                        else:
                            continue

                    for i in range(3, 203):
                        progress['value'] += 0.5
                        cell = ws.cell(row=i, column=1)
                        cell_val = cell.value
                        if cell_val:

                            if cell_val == 'Zákonné pojištění' or cell_val == 'Mzdové náklady':
                                break

                            if cell_val in employees_inter:
                                text_field.insert(tk.END, f"|- Vyplnuji {cell_val}: {employees_inter[cell_val][0]}\n")
                                logging.info(f"|- Vyplnuji {cell_val}: {employees_inter[cell_val][0]}")
                                ws.cell(row=i, column=month_col).value = employees_inter[cell_val][0]
                                ws.cell(row=i, column=month_col + fare_offset).value = employees_inter[cell_val][1]
                        root.update_idletasks()
                    logging.info("")

                wb.save(input_output['output_file_loc'])

    text_field.insert(tk.END, 'DONE')
    text_field.see('end')
    logging.info("DONE")


with open('structure.json', 'r', encoding='cp1250') as jdata:
    last_data = json.load(jdata)


def main_window(widget, width=0, height=0):
    screen_w, screen_h = widget.winfo_screenwidth(), widget.winfo_screenheight()

    left = int(screen_w / 2) - int(width / 2)
    top = int(screen_h / 2) - int(height / 2)

    if width and height:
        widget.geometry(f'{width}x{height}+{left}+{top}')
    else:
        widget.geometry('+%d+%d' % (500, 100))

    widget.resizable(0, 0)


root = tk.Tk()
root.grid_columnconfigure(0, weight=1)

main_window(root, 720, 610)
opts = {'padx': 10, 'sticky': 'WE', 'ipadx': 10, 'ipady': 10}

top_frame = ttk.Frame(root, height=180, style="GrooveBorder.TFrame")

top_frame.grid_columnconfigure(0, weight=1)
top_frame.grid_columnconfigure(1, weight=1)
top_frame.grid_columnconfigure(2, weight=6)

chosen_month = tk.StringVar()
choose_month = ttk.OptionMenu(top_frame, chosen_month, current_month, *months_cz)
choose_month.configure(width=8)

progress = ttk.Progressbar(top_frame, length=100, mode='determinate', orient='horizontal')

log_file_btn = ttk.Button(top_frame, width=1, text='Open log file', command=open_log_file)

banner = '''    _     __  __   _  _ 
   /_\   |  \/  | | \| |	.. Automatizace	..
  / _ \  | |\/| | | .` |	.. Mzdových		   ..
 /_/ \_\ |_|  |_| |_|\_|	.. Nákladů			   ..
============================================
'''
txt = ScrolledText(top_frame, width=2, height=5)
txt.insert(
    '1.0',
    banner
)

start_btn = ttk.Button(top_frame, width=1, text='Start', state='enabled',
                       command=lambda: amn(chosen_month.get(), last_data, txt))

btn_opts = {'sticky': 'we', 'padx': 5, 'pady': 5}

choose_month.grid(row=0, column=0, **btn_opts)
progress.grid(row=1, column=1, **btn_opts)
start_btn.grid(row=1, column=0, **btn_opts)
log_file_btn.grid(row=0, column=1, **btn_opts)
txt.grid(row=0, column=2, rowspan=2, columnspan=2, **btn_opts)

top_frame.grid(row=0, column=0, **opts, pady=10)

bottom_frame = ttk.Frame(root, height=400)
bottom_frame.grid(row=1, column=0, **opts)

tabs = ttk.Notebook(bottom_frame, width=700, height=400)
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
    # company_frame.grid_columnconfigure(4, weight=1)
    # company_frame.grid_columnconfigure(5, weight=1)

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
            up_btn_in,
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
            inter_btn_in,
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
        text=Path(file_paths['output_file_up']).name,
        width=20,
        command=lambda: set_datas(
            up_btn_out,
            tabs.select(),
            [('Excel sheet', '*.xls?')],
            'output_file_up'
        )
    )

    up_btn_out.grid(row=3, column=1)

    inter_label_out = ttk.Label(company_frame, text='interni')
    inter_label_out.grid(row=4, column=1)

    inter_btn_out = ttk.Button(
        company_frame,
        text=Path(file_paths['output_file_loc']).name,
        width=20,
        command=lambda: set_datas(
            inter_btn_out,
            tabs.select(),
            [('Spreadsheets', '*.xlsx')],
            'output_file_loc'
        )
    )

    inter_btn_out.grid(row=5, column=1)

    data_label = ttk.Label(company_frame, text='Data')
    data_label.grid(row=0, column=2)

    data_btn = ttk.Button(
        company_frame,
        text=Path(file_paths['input_data']).name,
        width=20,
        command=lambda: set_datas(
            data_btn,
            tabs.select(),
            [('CSV', '*.csv')],
            'input_data'
        )
    )

    data_btn.grid(row=0, column=3, columnspan=3)

    ins_groups_label = ttk.Label(company_frame, text='POJISTOVNY')
    ins_groups_label.grid(row=1, column=2)

    tree_label = ttk.Label(company_frame, text='kody zdravotnich pojistoven v ucto2000')
    tree_label.grid(row=2, column=2, columnspan=3)
    tree_cols = ('ucto_code', 'ins_code', 'ins_name')
    tree = ttk.Treeview(company_frame, columns=tree_cols, show='headings')
    tree.column('ucto_code', width=70)
    tree.column('ins_code', width=100)
    tree.heading('ucto_code', text='Ucto2000')
    tree.heading('ins_code', text='kod pojistovny')
    tree.heading('ins_name', text='nazev pojistovny')

    load_ins_codes(company_name, tree)

    tree.bind('<<TreeviewSelect>>', lambda event='<<TreeviewSelect>>', my_tree=tree: item_selected(event, my_tree))
    tree.grid(row=3, column=2, columnspan=3, rowspan=3)
    scrollbar = ttk.Scrollbar(company_frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscroll=scrollbar.set, height=5)

    code_ucto_btn = ttk.Entry(company_frame)
    code_ucto_btn.grid(row=7, column=2)

    ins_code_btn = ttk.Entry(company_frame)
    ins_code_btn.grid(row=7, column=3)

    ins_name_btn = ttk.Entry(company_frame)
    ins_name_btn.grid(row=7, column=4)

    add_button = ttk.Button(company_frame, text="Add new", command=lambda: add_new(tabs.select()))
    add_button.grid(row=8, column=2)

    update = ttk.Button(company_frame, text="Update selected",
                        command=lambda my_tree=tree: update_ins(tabs.select(), my_tree))
    update.grid(row=8, column=3)

    delete_btn = ttk.Button(company_frame, text="Delete record",
                            command=lambda my_tree=tree: delete_record(tabs.select(), my_tree))
    delete_btn.grid(row=8, column=4)

    for w in company_frame.winfo_children():
        w.grid(padx=10, pady=10, sticky='NWSE')
    scrollbar.grid(row=3, column=5, sticky='NS', rowspan=3)

    tabs.add(company_frame, text=company_name)
tabs.add(ttk.Frame(tabs), text='+')

root.mainloop()
