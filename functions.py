from pathlib import Path
import json
import csv
import tkinter as tk
import subprocess
from tkinter import simpledialog
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.messagebox import askyesno
from os import remove

HOME = Path.home() / 'amn'
STRUCTURE = HOME / 'structure.json'
if not Path(STRUCTURE).exists():
    STRUCTURE = 'structure.json'


def check_new_ppl(input_file):
    with open(input_file, 'r', encoding='cp1250') as pracov:
        new_or_dead_p = {}
        new_or_dead_loc = {}
        dreader = csv.DictReader(pracov)
        to_add_mil = ['VstupDoZam', 'UkonceniZam', 'DuchOd']

        for row in dreader:
            name = row['Jmeno30'][1:-1]
            rodcis = row['RodCislo'].replace('/', '').replace('"', '').replace('\'', '').replace('\n', '')
            new_or_dead_p.setdefault(rodcis, {})
            new_or_dead_loc.setdefault(name, {})

            for d in to_add_mil:
                if row[d]:
                    if int(row[d][-3:-1]) <= 22:
                        new_or_dead_p[rodcis][d] = row[d][1:-3] + '20' + row[d][-3:-1]
                    else:
                        new_or_dead_p[rodcis][d] = row[d][1:-3] + '19' + row[d][-3:-1]
                else:
                    new_or_dead_p[rodcis][d] = ''

            if 'invalidní 3.stup' in row['TypDuch']:
                pension_type = 'TZP'
            elif 'invalidní 1.nebo' in row['TypDuch']:
                pension_type = 'OZP12'
            else:
                pension_type = ''

            new_or_dead_p[rodcis]['TypDuch'] = pension_type
                   
            new_or_dead_loc[name]['Kod'] = row['Kod'][1:-1]


    return new_or_dead_p, new_or_dead_loc


def prepare_input(input_file, c_name):

    up_table = {}
    inter_table = {}
    ins_codes = {}

    if Path(f'insurance_codes_{c_name}.json').exists():
        with open(f'insurance_codes_{c_name}.json', 'r') as ins_file:
            ins_codes = json.load(ins_file)

    with open(input_file, 'r', encoding='cp1250') as empl_data_csv:
        empl_dict_reader = csv.DictReader(empl_data_csv)

        for row in empl_dict_reader:

            rodcis = row['RodCislo'][1:-1].replace('/', '')
            full_name = row['JmenoS'][1:-1].split(' ')

            # Name might include title, merge first name with title
            if len(full_name) > 2:
                fname = " ".join(full_name[1:])
            else:
                fname = full_name[1]

            fname = fname.strip()
            lname = full_name[0].strip()

            ins = row['CisPoj'][1:-1]
            cat = row['Kat'][1:-1]
            fare = 0

            for fare_cost in [row['Davky1'], row['Davky2']]:
                if fare_cost:
                    fare += int(fare_cost)

            exp = 0
            total_exp = 0
            
            if cat == 'INV':
                for cost in [row['HrubaMzda'], row['Zamest'], row['iNemoc']]:
                    if cost:
                        exp += int(cost)
                total_exp = exp
            else:
                for cost in [row['HrubaMzda'], row['Zamest']]:
                    if cost:
                        exp += int(cost)
                if row['iNemoc']:
                    total_exp = exp + int(row['iNemoc'])
                else:
                    total_exp = exp

            if ins in ins_codes:
                ins_group_code = ins_codes[ins][0]
            else:
                ins_group_code = 999
                
            up_table.setdefault(rodcis, {'first name': fname, 'last name': lname, 'ins code': ins_group_code, 'cat': cat, 'payment expenses': total_exp})
            full_name = lname + ' ' + fname
            full_name = full_name[:20]
            inter_table.setdefault(full_name, (total_exp, fare))

    return up_table, inter_table


def main_window(widget, width=0, height=0):
    screen_w, screen_h = widget.winfo_screenwidth(), widget.winfo_screenheight()

    left = int(screen_w / 2) - int(width / 2)
    top = int(screen_h / 2) - int(height / 2)

    if width and height:
        widget.geometry(f'{width}x{height}+{left}+{top}')
    else:
        widget.geometry('+%d+%d' % (500, 100))

    widget.resizable(0, 0)


def show_banner(txt):
    txt.delete('1.0', tk.END)
    banner = '''    _     __  __   _  _ 
   /_\   |  \/  | | \| |			   |       Automatizace	
  / _ \  | |\/| | | .` |	     |       Mzdových		   
 /_/ \_\ |_|  |_| |_|\_|			   |       Nákladů			   
========================================================= '''
    txt.insert('1.0', banner)
    get_help(txt)


def get_help(txt):
    with open('help.txt', 'r') as helpfile:
        h = helpfile.read()
    txt.insert(tk.END, h)


def load_ins_codes(c_name, my_tree):
    my_tree.delete(*my_tree.get_children())
    ins_codes_file = f'insurance_codes_{c_name}.json'
    if Path(ins_codes_file).exists():
        ins_codes = json.load(open(ins_codes_file))
        for ins_code, data in ins_codes.items():
            my_tree.insert('', tk.END, values=[ins_code, data[0], data[1]])


def add_new(root, c_name, text_win, tab):

    text_win.delete('1.0', tk.END)
    my_tree = root.nametowidget(tab + '.!treeview')
    # c_name = tabs.tab(tab)['text']

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
            text_win.insert('1.0', 'Tento kod uz existuje.')
            print('This code is already used')
    else:
        text_win.insert('1.0', 'Vyplnte vsechna pole.')
        print('Fill out all fields')

    save_ins(c_name, my_tree)


def update_ins(window, c_name, tab, my_tree):
    selected = my_tree.focus()
    ucto_code_entry = str(window.nametowidget(tab + '.!entry').get())
    ins_code_entry = str(window.nametowidget(tab + '.!entry2').get())
    ins_name_entry = window.nametowidget(tab + '.!entry3').get()

    values = (ucto_code_entry, ins_code_entry, ins_name_entry)
    my_tree.item(selected, text="", values=values)

    save_ins(c_name, my_tree)


def delete_record(c_name, my_tree):
    selected = my_tree.focus()
    my_tree.delete(selected)
    save_ins(c_name, my_tree)


def save_ins(c_name, my_tree):
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
        outfile.write(json.dumps(ins_data, indent=4))
        # json.dump(ins_data, outfile)


def rename_tab(window, companies, last_data, tab, tabs):
    c_name = tabs.tab(tab)['text']
    btn = window.nametowidget(tab + '.!checkbutton')
    new_name = simpledialog.askstring("Prejmenovat", "Novy nazev", parent=window)
    
    confirm = askyesno(title="POZOR", message="Opravdu chcete prejmenovat firmu?")
    
    if confirm:

        if c_name in companies:
            index = companies.index(c_name)
            companies.pop(index)
            companies.append(new_name)

        tabs.tab(tab, text=new_name)
        new_data = {new_name if k == c_name else k:v for k, v in last_data.items()}
        btn['text'] = new_name

        with open(STRUCTURE, 'w') as outfile:
            outfile.write(json.dumps(new_data, indent=4))

        if Path(f'insurance_codes_{c_name}.json').exists():
            Path(f'insurance_codes_{c_name}.json').rename(f'insurance_codes_{new_name}.json')


def delete_tab(last_data, companies, tabs):

    confirm = askyesno(title="POZOR", message="Chcete smazat firmu?")
    
    if confirm:
        c_name = tabs.tab(tabs.select())['text']
        ins_file = Path(f'insurance_codes_{c_name}.json')

        if c_name in last_data:
            del last_data[c_name]
            with open(STRUCTURE, 'w') as outfile:
                outfile.write(json.dumps(last_data, indent=4))

        if c_name in companies:
            companies.remove(c_name)

        if ins_file.exists():
            print('deleting insurance file.')
            remove(ins_file)
        else:
            print('No insurance file found.')

        tabs.forget(tabs.select())


def open_output(c_name, last_data):
    output_folder = last_data[c_name]['output']
    subprocess.Popen(f'explorer {output_folder}')


def item_selected(window, event, my_tree):
    c_frame = my_tree.winfo_parent()
    ucto_code_entry = window.nametowidget(c_frame + '.!entry')
    ins_code_entry = window.nametowidget(c_frame + '.!entry2')
    ins_name_entry = window.nametowidget(c_frame + '.!entry3')

    ucto_code_entry.delete(0, tk.END)
    ins_code_entry.delete(0, tk.END)
    ins_name_entry.delete(0, tk.END)

    selected_item = my_tree.focus()
    row = my_tree.item(selected_item)
    record = row['values']

    ucto_code_entry.insert(0, str(record[0]).zfill(5))
    ins_code_entry.insert(0, record[1])
    ins_name_entry.insert(0, record[2])


def activate_tab(window, tab, act, start_btn, c_name, companies):
    state = act.get()
    # c_name = tabs.tab(tab)['text']
    buttons = ['.!button',
               '.!button2',
               '.!button5',
               '.!button6',
               '.!button7',
               '.!button8',
               '.!button9',
               '.!button10',
               '.!entry',
               '.!entry2',
               '.!entry3']

    if state == '0':
        for b in buttons:
            window.nametowidget(tab + b)['state'] = 'disabled'
        if c_name in companies:
            companies.remove(c_name)
    else:
        for b in buttons:
            window.nametowidget(tab + b)['state'] = 'enabled'
        if c_name not in companies:
            companies.append(c_name)

    if len(companies) > 0:
        start_btn['state'] = 'enabled'
    else:
        start_btn['state'] = 'disabled'


def set_dir(window, c_name, btn, comp, last_data):
    last_dir = Path(last_data[c_name]['output'])
    if last_dir.exists():
        initdir = last_dir
    else:
        initdir = Path.home()
    
    dir_name = askdirectory(title='Choose output folder', initialdir=initdir)

    if dir_name:
        window.nametowidget(comp + '.' + btn.winfo_name())["text"] = str(Path(dir_name).name)
        last_data[c_name]['output'] = str(Path(dir_name))
        with open(STRUCTURE, 'w', encoding='cp1250') as input_file:
            input_file.write(json.dumps(last_data, indent=4))


def set_datas(window, c_name, last_data, btn, comp, filetypes, key_name):

    last_dir = Path(last_data[c_name][key_name]).parent
    if last_dir.exists():
        initdir = last_dir
    else:
        initdir = Path.home()
        

    filename = askopenfilename(
        title='Open a file',
        initialdir=initdir,
        filetypes=filetypes
    )

    if filename:
        file = Path(filename).name
        window.nametowidget(comp + '.' + btn.winfo_name())['text'] = file

        last_data[c_name][key_name] = str(Path(filename))
        with open(STRUCTURE, 'w', encoding='cp1250') as input_file:
            input_file.write(json.dumps(last_data, indent=4))
