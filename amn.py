# TOOD:
# 1) a) choose month; input num <= 12; month = input % 3 for UP
#    b) month = months_cs[input - 1] for mzdove naklady
# 2) mzdove naklady
# 3) choose bereko or fiala from filename of csv file; bereko.csv or fiala.csv
# 4) update both UP and mzdove naklady in one run
# 5) update both fiala and bereko in one run
# --------------------------------------------------------------------------------
# maybe not possible:
# 6) get status of ozp, add to exported csv
# 7) get placement code for mzdove naklady to insert new employee to correct sheet


import logging
import csv
import openpyxl
from pathlib import Path
from openpyxl.utils import column_index_from_string

logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

INPUT_FOLDER = Path.cwd() / 'INPUT'
OUTPUT_FOLDER = Path.cwd() / 'OUTPUT'
SRC_FOLDER = Path.cwd() / 'SRC'

empl_data_csv_file = INPUT_FOLDER / 'I.CSV'
with open(empl_data_csv_file, 'r') as empl_data_csv:
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
    # logging.info(i)
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

    employees_inter.setdefault(lname + ' ' + fname, total_expences)

# logging.info(employees_up)
# logging.info(employees_inter)

# open mzdy UP table, go through each name in data and check if it is in table
wb = openpyxl.load_workbook(SRC_FOLDER / 'jmenny_seznam_2021_10_01 Bereko.xlsx')
ws = wb['2) jmenný seznam']
col_letter_id = column_index_from_string('D')
col_letter_pay = column_index_from_string('K')

ids_in_xlsx = {}
last_row = 13

for i in range(10):
    row = i + 13
    id_of_person = ws.cell(row=row, column=col_letter_id).value
    # make list of ids in xlsx file
    if id_of_person:
        ids_in_xlsx.setdefault(id_of_person, {'row': row})
        last_row = row

in_xlsx_only = {}
in_xlsx_and_csv_data = {}
in_csv_only = {}

for i in employees_up:
    if i not in ids_in_xlsx:
        # new employee or asistenti sheet
        in_csv_only[i] = employees_up[i]
    else:
        # grab cell coord and add it to dict in_xlsx_and_csv_data
        in_xlsx_and_csv_data[i] = employees_up[i]
        in_xlsx_and_csv_data[i]['row'] = ids_in_xlsx[i]['row']

# person located in xlsx and is in csv
# insert payment expences to each employee that is in table and is
for person in in_xlsx_and_csv_data:
    # logging.info(in_xlsx_and_csv_data[person]['row'])
    row = in_xlsx_and_csv_data[person]['row']
    ws.cell(row=row, column=column_index_from_string('K')).value = in_xlsx_and_csv_data[person]['payment expences']


# person located in xlsx and is NOT in the csv
for i in ids_in_xlsx:
    if i not in employees_up:
        ws.cell(row=ids_in_xlsx[i]['row'], column=column_index_from_string('J')).value = ''


# person NOT located in xlsx and is in the csv => new employee
if len(in_csv_only) > 0:
    ws = wb['3) nákl. prov. z. a prac. a.']
    for person in in_csv_only:
        # logging.info(in_csv_only[person])
        person_id = person[:6] + '/' + person[6:]
        located = False
        last_row_asist = 14
        for i in range(last_row_asist, 20):
            rc = ws.cell(row=i, column=column_index_from_string('D')).value
            if rc == person_id:
                ws.cell(row=i, column=column_index_from_string('I')).value = in_csv_only[person]['payment expences']
                located = True
                last_row_asist = i
        if not located:
            if in_csv_only[person]['cat'] == 'INV':
                ws = wb['2) jmenný seznam']
                ws.cell(row=last_row+1, column=column_index_from_string('B')).value = in_csv_only[person]['last name']
                ws.cell(row=last_row+1, column=column_index_from_string('C')).value = in_csv_only[person]['first name']
                ws.cell(row=last_row+1, column=column_index_from_string('D')).value = person
                ws.cell(row=last_row+1, column=column_index_from_string('G')).value = in_csv_only[person]['ins code']
                ws.cell(row=last_row+1, column=column_index_from_string('K')).value = in_csv_only[person]['payment expences']
            else:
                ws = wb['3) nákl. prov. z. a prac. a.']
                ws.cell(row=last_row_asist+1, column=column_index_from_string('B')).value = in_csv_only[person]['last name']
                ws.cell(row=last_row_asist+1, column=column_index_from_string('C')).value = in_csv_only[person]['first name']
                ws.cell(row=last_row_asist+1, column=column_index_from_string('D')).value = person_id
                ws.cell(row=last_row_asist+1, column=column_index_from_string('E')).value = '-\'\'-'
                ws.cell(row=last_row_asist+1, column=column_index_from_string('F')).value = 'PA'
                ws.cell(row=last_row_asist+1, column=column_index_from_string('G')).value = '100%'
                ws.cell(row=last_row_asist+1, column=column_index_from_string('I')).value = in_csv_only[person]['payment expences']

# save as tempcopy-up-fiala.xlsx
wb.save(OUTPUT_FOLDER / 'temp.xlsx')
# open document for visual check

# open mzdove naklady fiala, go through each sheet and check if name is in data
# if found, add payment expences to correct month
# if not found, add name to first column and fill in payment expences

# repeat for bereko
