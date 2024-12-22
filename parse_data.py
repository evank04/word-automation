import openpyxl
import csv
import re


def format_date_time(text):
    return text.replace(" ", '\n').replace("'", "")


def filter_string(text):
    return text.replace(' ', '').replace("'", "")


# returns a list of the test result and date-time for each table
def fpc_read_table(file) -> tuple:
    contents = csv.reader(file)
    table = [line for line in contents]
    datas = table[2:4]
    test_result = [
        datas[0][3], datas[0][4],
        datas[1][3], datas[1][4]
    ]

    date_time = [format_date_time(datas[0][7]), format_date_time(datas[1][7])]
    filtered_test_result = [filter_string(result) for result in test_result]
    return filtered_test_result, date_time


def format_string(text):
    patterns = {
        'FL': "FL\\(-?[0-9]+,-?[0-9]+\\)",
        'FR': "FR\\(-?[0-9]+,-?[0-9]+\\)",
        'RL': "RL\\(-?[0-9]+,-?[0-9]+\\)",
        'RR': "RR\\(-?[0-9]+,-?[0-9]+\\)"
    }

    formatted_data = []

    for label, pattern in patterns.items():
        matches = re.findall(pattern, text)
        formatted_data.append(matches[0]) if matches else formatted_data.append(f"{label}(,)")

    return "\n".join(formatted_data)


wb = openpyxl.load_workbook(f'C:\\Users\\Evan\\Desktop\\proj\\Base Template\\datas2.xlsx')

table_number = 1

formatted_test_result = []
fpc_date_time = []
while table_number < 7:
    with open(f'C:\\Users\\Evan\\Desktop\\proj\\ocr\\Binder1\\table-{table_number}.csv') as file:
        lst, timing = fpc_read_table(file)
        for i in lst:
            formatted_test_result.append(format_string(i))
        for time in timing:
            fpc_date_time.append(time)
    table_number += 1

FPC = wb['FPC']

fps_coordinates = [
    ('D5', 'E5'), ('D6', 'E6'),
    ('D11', 'E11'), ('D12', 'E12'),
    ('D17', 'E17'), ('D18', 'E18'),
    ('D23', 'E23'), ('D24', 'E24'),
    ('D29', 'E29'), ('D30', 'E30'),
    ('D35', 'E35'), ('D36', 'E36')
]

fps_time_coordinates = {'H': [5, 6, 11, 12, 17, 18, 23, 24, 29, 30, 35, 36]}

for i, (d, e) in enumerate(fps_coordinates):
    FPC[d] = formatted_test_result[2 * i]
    FPC[e] = formatted_test_result[2 * i + 1]

for column, rows in fps_time_coordinates.items():
    for i, row in enumerate(rows):
        FPC[column + str(row)] = fpc_date_time[i]


def ofpc_read_table(file):
    contents = csv.reader(file)
    table = [line for line in contents]
    datas = table[2:4]
    test_result = [datas[0][2].replace("'", ""), datas[1][2].replace("'", "")]
    date_time = [datas[0][3].replace("'", ""), datas[1][3].replace("'", "")]
    return test_result, date_time


ofpc_test_result = []
ofpc_date_time = []
ofpc_table_number = 1
while ofpc_table_number < 15:
    with open(f'C:\\Users\\Evan\\Desktop\\proj\\ocr\\Binder2\\table-{ofpc_table_number}.csv') as file:
        x, y = ofpc_read_table(file)
        for i in x:
            ofpc_test_result.append(i)
        for i in y:
            ofpc_date_time.append(i)
    ofpc_table_number += 1

ofpc_rows = [5, 6, 11, 12, 17, 18, 23, 24, 29, 30, 35, 36, 41, 42]

ofpc_coordinates = {
    'C': ofpc_rows,
    'F': ofpc_rows
}

ofpc_time_coordinates = {
    'D': ofpc_rows,
    'H': ofpc_rows
}

OFPC = wb['OFPC']

ofpc_helper1 = 0
for column, rows in ofpc_coordinates.items():
    for row in rows:
        OFPC[column + str(row)] = ofpc_test_result[ofpc_helper1]
        ofpc_helper1 += 1

ofpc_helper2 = 0
for column, rows in ofpc_time_coordinates.items():
    for row in rows:
        OFPC[column + str(row)] = ofpc_date_time[ofpc_helper2]
        ofpc_helper2 += 1


def read_last_tables(file):
    global last_table_num
    contents = csv.reader(file)
    table = [line for line in contents]
    if last_table_num == 1:
        cdods_result = table[1][1].replace("'", "")
        return cdods_result
    elif last_table_num == 2:
        datas = table[1:4]
        hobds_result = [datas[0][4].replace("'", ""), datas[1][4].replace("'", ""), datas[2][4].replace("'", "")]
        return hobds_result
    else:
        datas = table[2:6]
        systems_result = [datas[0][5].replace("'", ""), datas[1][5].replace("'", ""), datas[2][5].replace("'", ""),
                          datas[3][5].replace("'", "")]
        return systems_result


cdods_coordinate = "D5"
hobds_coordinate = {"F": [11, 12, 13]}
systems_coordinates = {
    "F1": [6, 7, 8, 9],
    "L1": [6, 7, 8, 9],
    "F2": [15, 16, 17, 18],
    "L2": [15, 16, 17, 18],
    "F3": [24, 25, 26, 27],
    "L3": [24, 25, 26, 27],
}

SSC = wb["SSC"]
CLPS = wb["CLPS"]
CLODS = wb["CLODS"]
systems_result = []
last_table_num = 1
while last_table_num < 15:
    with open(f'C:\\Users\\Evan\\Desktop\\proj\\ocr\\Binder3\\table-{last_table_num}.csv') as file:
        x = read_last_tables(file)
        if last_table_num == 1:
            SSC[cdods_coordinate] = x
            last_table_num += 1
        elif last_table_num == 2:
            for column, rows in hobds_coordinate.items():
                for i, row in enumerate(rows):
                    SSC[column + str(row)] = x[i]
            last_table_num += 1
        else:
            for result in x:
                systems_result.append(result)
            last_table_num += 1


helper3 = 0
table_count = 1

while helper3 < len(systems_result):
    if table_count <= 8:
        for column, rows in systems_coordinates.items():
            for row in rows:
                CLPS[column[0] + str(row)] = systems_result[helper3]
                table_count += 1
                helper3 += 1
    elif table_count > 8:
        for column, rows in systems_coordinates.items():
            for row in rows:
                CLODS[column[0] + str(row)] = systems_result[helper3]
                table_count += 1
                helper3 += 1


wb.save(f'C:\\Users\\Evan\\Desktop\\proj\\Excel Data\\ocr_latest.xlsx')
