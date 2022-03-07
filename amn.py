# amn.py
# -------------+
# Automatizace |
# Mzdovych     |
# Nakladu      |
# -------------+
# v1
# author: Zdenek Frydryn
# created for: Bereko s.r.o., Drogerie Fiala s.r.o.,
# Description: Automatizace vyplnovani mezd do xlsx tabulek pro urad prace.

# TOOD:
# 1) Check for quartal
# 2) mzdove naklady
# 3) choose bereko or fiala from filename of csv file; bereko.csv or fiala.csv
# 4) update both UP and mzdove naklady in one run
# 5) update both fiala and bereko in one run
# --------------------------------------------------------------------------------
# maybe not possible:
# 6) get status of ozp, add to exported csv


import logging
import csv
import openpyxl
import pyinputplus as pyin
from pathlib import Path
from openpyxl.utils import column_index_from_string
from months_cz import months_cz

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s',
                    datefmt='%Y-%m-%d %H:%M:%S')

INPUT_FOLDER = Path.cwd() / 'INPUT'
OUTPUT_FOLDER = Path.cwd() / 'OUTPUT'
SRC_FOLDER = Path.cwd() / 'SRC'

COMPANIES = {
    'Fiala': {
        'input_data': INPUT_FOLDER / 'fiala.CSV',
        'src_file': SRC_FOLDER / 'jmenny_seznam_2021_10_01 Fiala.xlsx',
        'output_file': OUTPUT_FOLDER / 'temp-fiala.xlsx'
    },
    'Bereko': {
        'input_data': INPUT_FOLDER / 'bereko.CSV',
        'src_file': SRC_FOLDER / 'jmenny_seznam_2021_10_01 Bereko.xlsx',
        'output_file': OUTPUT_FOLDER / 'temp-bereko.xlsx'
    }
}

MONTH = pyin.inputMenu(months_cz, numbered=True, prompt="Vyberte mesic:\n")
MONTH = months_cz.index(MONTH) + 1

for company_name, input_output in COMPANIES.items():
    logging.info(f'Vyplnuji {company_name}')

    with open(input_output['input_data'], 'r') as empl_data_csv:
        empl_data_reader = csv.reader(empl_data_csv)
        empl_data = list(empl_data_reader)
        empl_data = empl_data[1:]

    employees_up = {}
    employees_inter = {}
    insurance_codes = {'00903': 111, '00015': 201}

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
        if ins in insurance_codes:
            ins_group_code = insurance_codes[ins]
        else:
            ins_group_code = 999

        # get category of employment contract
        category = i[3][1:-1]

        # calculate salary
        total_expences = 0
        for exp in i[4:]:
            try:
                total_expences += int(exp)
            except ValueError:
                continue

        employees_up.setdefault(id_num, {'first name': fname,
                                         'last name': lname,
                                         'ins code': ins_group_code,
                                         'cat': category,
                                         'payment expences': total_expences
                                         })

        # employees_inter.setdefault(lname + ' ' + fname, total_expences)

    # logging.info(employees_inter)

    # open mzdy UP table, go through each name in data and check if it is in table
    wb = openpyxl.load_workbook(SRC_FOLDER / input_output['src_file'])
    sheets = ['2) jmenný seznam', '3) nákl. prov. z. a prac. a.']
    col_letter_id = column_index_from_string('D')
    if MONTH % 3 == 1:
        col_letter_pay = {sheets[0]: 11, sheets[1]: 9}
    elif MONTH % 3 == 2:
        col_letter_pay = {sheets[0]: 16, sheets[1]: 10}
    else:
        col_letter_pay = {sheets[0]: 21, sheets[1]: 11}

    ids_in_xlsx = {}
    last_rows = []

    for sheet in sheets:
        ws = wb[sheet]
        last_row = 13
        ids_in_xlsx.setdefault(sheet, {})

        for i in range(10):
            row = i + 13
            id_of_person = ws.cell(row=row, column=col_letter_id).value
            # make list of ids in xlsx file
            if id_of_person:
                ids_in_xlsx[sheet].setdefault(id_of_person.replace('/', ''), row)
                last_row = row

        last_rows.append(last_row)

    for emp_id, emp_data in employees_up.items():
        found = False
        for sheet_name, sheet_data in ids_in_xlsx.items():
            if emp_id in sheet_data:
                found = True
        if not found:
            logging.info(f"Novy zamestnanec {emp_data['first name']} {emp_data['last name']}")
            if emp_data['cat'] == 'INV':
                logging.info(f"Pridavam do listu {sheets[0]}")
                ws = wb[sheets[0]]
                ws.cell(row=last_rows[0] + 1, column=2).value = emp_data['last name']
                ws.cell(row=last_rows[0] + 1, column=3).value = emp_data['first name']
                ws.cell(row=last_rows[0] + 1, column=4).value = emp_id
                ws.cell(row=last_rows[0] + 1, column=7).value = emp_data['ins code']
                ws.cell(row=last_rows[0] + 1, column=col_letter_pay[ws.title]).value = emp_data['payment expences']
                last_rows[0] += 1

            else:
                logging.info(f"Pridavam do listu {sheets[1]}")
                ws = wb[sheets[1]]
                ws.cell(row=last_rows[1] + 1, column=2).value = emp_data['last name']
                ws.cell(row=last_rows[1] + 1, column=3).value = emp_data['first name']
                ws.cell(row=last_rows[1] + 1, column=4).value = emp_id[:6] + '/' + emp_id[6:]
                ws.cell(row=last_rows[1] + 1, column=col_letter_pay[ws.title]).value = emp_data['payment expences']
                ws.cell(row=last_rows[1] + 1, column=5).value = '-\'\'-'
                ws.cell(row=last_rows[1] + 1, column=6).value = 'PA'
                ws.cell(row=last_rows[1] + 1, column=7).value = '100%'
                last_rows[1] += 1

    logging.info("-" * 32)

    for sheet_name, data in ids_in_xlsx.items():
        ws = wb[sheet_name]

        for id_num, row_num in data.items():
            if id_num in employees_up:
                logging.info(
                    f"Vyplnuji vyplatu pro {employees_up[id_num]['first name']} {employees_up[id_num]['last name']} v listu {sheet_name}")
                # zadat plat za tento mesic
                ws.cell(row=row_num, column=col_letter_pay[sheet_name]).value = employees_up[id_num]['payment expences']

            else:
                logging.info(f"Zamestnanec {id_num} v listu {sheet_name} nema vyplnenou vyplatu. Mazu pole status")
                # smazat v xlsx status pro tento mesic
                if sheet_name == sheets[0]:
                    ws.cell(row=row_num, column=col_letter_pay[sheet_name] - 1).value = ''

    wb.save(input_output['output_file'])
    logging.info(f"{company_name} vyplneno.\n")
    # open document for visual check

    # open mzdove naklady fiala, go through each sheet and check if name is in data
    # if found, add payment expences to correct month
    # if not found, add name to first column and fill in payment expences

    # repeat for bereko
logging.info("Done")
