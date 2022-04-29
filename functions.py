from pathlib import Path
import json
import tkinter as tk


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


def activate_tab(tab, act):
    state = act.get()
    btn = root.nametowidget(tab + '.!checkbutton')
    buttons = ['.!button',
               '.!button2',
               '.!button3',
               '.!button4',
               '.!button5',
               '.!button6',
               '.!button7',
               '.!entry',
               '.!entry2',
               '.!entry3']

    if state == '0':
        for b in buttons:
            root.nametowidget(tab + b)['state'] = 'disabled'
    else:
        for b in buttons:
            root.nametowidget(tab + b)['state'] = 'enabled'


def set_dir(btn, comp):
    c_name = tabs.tab(comp)["text"]
    dir_name = askdirectory(title='Choose output folder', initialdir='.')
    if dir_name:
        root.nametowidget(comp + '.' + btn.winfo_name())["text"] = str(Path(dir_name).name)
        last_data[c_name]['output'] = str(Path(dir_name))
        with open('structure.json', 'w', encoding='cp1250') as input_file:
            input_file.write(json.dumps(last_data))


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
