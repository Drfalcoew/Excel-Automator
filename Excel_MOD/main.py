# Excel Organizer
# Drew Foster, Python, 11/21/2021
import re

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, PatternFill, Font, Border

my_wb = load_workbook('spreadsheet_template_.xlsx')
my_ws = my_wb.active
end_row = my_ws.max_row
end_col = my_ws.max_column
my_ws.title = 'OG'
col = get_column_letter(end_col)
masterpack_fill = PatternFill(start_color='8497b0',
                              end_color='8497b0',
                              fill_type='solid')


def filter_red():
    red = "FFFF0000"
    black = "FF424649"
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        if my_ws['A' + str(temp)].font.color.rgb == red:
            my_ws.delete_rows(temp)
            print('A{} font: red. Deleting'.format(temp))
        elif my_ws['A' + str(temp)].font.color.rgb == black:
            print('A{} font: black. Continuing'.format(temp))
        temp -= 1
    print("...Finished")


def filter_scac():
    fedex = "FXGR"
    ups = "UP"
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        scac_value = my_ws['AL' + str(temp)].value
        if scac_value == "" or scac_value is None:
            print('A{} SCAC: {}. Continuing'.format(temp, scac_value))
        elif scac_value == fedex or scac_value[0:2] == ups:
            print('A{} SCAC: {}. Continuing'.format(temp, scac_value))
        else:
            my_ws.delete_rows(temp)
            print('A{} SCAC: {}. Deleting'.format(temp, scac_value))
        temp -= 1
    print("...Finished")


def filter_freight():
    tp = "TP"
    p = "P"
    pi = "PI"
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        freight_value = my_ws['Y' + str(temp)].value
        if freight_value == "" or freight_value is None:
            print('A{} FREIGHT: {}. Continuing'.format(temp, freight_value))
        elif freight_value == tp or freight_value == p or freight_value == pi:
            print('A{} FREIGHT: {}. Continuing'.format(temp, freight_value))
        else:
            my_ws.delete_rows(temp)
            print('A{} FREIGHT: {}. Deleting'.format(temp, freight_value))
        temp -= 1
    print("...Finished")


def filter_customer_name():
    dunhams = "DUNHAMS SPORTS EDI"
    bj = "BJ'S WHOLESALE CLUB"
    walmart = "WALMART.COM EDI INVENTORY"
    scheels = "SCHEELS EDI"
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        customer_name = my_ws['AP' + str(temp)].value
        if customer_name == "" or customer_name is None:
            print('AP{} Customer Name: {}. Continuing'.format(temp, customer_name))
        elif customer_name == dunhams or customer_name == bj or customer_name == walmart or customer_name == scheels:
            print('AP{} Customer Name: {}. Deleting'.format(temp, customer_name))
            my_ws.delete_rows(temp)
        else:
            print('A{} Customer Name: {}. Continuing'.format(temp, customer_name))
        temp -= 1
    print("...Finished")


def filter_spec():
    pilot = "SHIP VIA PILOT"
    esfww = "SHIP VIA EFWW-ESTES FORWARDING WW"
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        spec = my_ws['AH' + str(temp)].value
        if spec == "" or spec is None:
            print('AH{} SPEC INSTR: {}. Continuing'.format(temp, spec))
        elif spec == esfww or spec == pilot:
            print('AH{} SPEC INSTR: {}. Deleting'.format(temp, spec))
            my_ws.delete_rows(temp)
        else:
            print('AH{} SPEC INSTR: {}. Continuing'.format(temp, spec))
        temp -= 1
    print("...Finished")


def filter_zero_weights():
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        weight = my_ws['G' + str(temp)].value
        _master_p = my_ws['H{0}'.format(str(temp))].value
        if weight == 0 or weight == 0.0 or weight == .0:
            print('G{} Weight: {}. Deleting'.format(temp, weight))
            my_ws.delete_rows(temp)
        else:
            print('G{} Weight: {}. Continuing'.format(temp, weight))

        if _master_p == 0:
            my_ws['H{0}'.format(str(temp))] = 1
        temp -= 1
    print("...Finished")


def filter_char_osadx():
    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting... (removing non-numerics from phone #s)")
    for row in range(1, _end_row):
        osadx = my_ws['M' + str(temp)].value
        if osadx != "" or osadx is not None:
            my_ws['M' + str(temp)] = re.sub("[^0-9]", "", str(osadx))
        temp -= 1
    print("...Finished")


def duplicate_cartons():
    value_list = []
    _end_row = my_ws.max_row
    temp = _end_row - 1
    print("Starting... (duplicating 2x)")
    for row in range(1, _end_row - 1):
        open_q = my_ws['F' + str(temp)].value
        master_p = my_ws['H{0}'.format(str(temp))].value
        if open_q > master_p == 1:
            print(f"Row {temp} duplicating")
            r = int(open_q / master_p) - 1
            print(r)
            my_ws['F' + str(temp)] = 1
            for dup in range(r):
                value_list.clear()
                for i in range(1, 43):
                    value_list.append(my_ws[str(get_column_letter(i)) + str(temp)].value)
                print(value_list)
                my_ws.append(value_list)
            print(value_list)
        temp -= 1
    print("...Finished")


def masterpacks():
    _end_row = my_ws.max_row
    temp = _end_row - 1
    value_list = []

    print("Starting... (Checking for MPs)")
    for _ in range(1, _end_row - 1):
        open_q = my_ws['F' + str(temp)].value
        master_p = my_ws['H{0}'.format(str(temp))].value
        weight = my_ws['G' + str(temp)].value

        if open_q is not None:
            if open_q > 1 and master_p != 1:
                print(f"Row {temp} is a masterpack")
                my_ws['G' + str(temp)].fill = masterpack_fill  # filling OG cell range with blue color
                if open_q > master_p:
                    # change original cell
                    my_ws['F' + str(temp)] = my_ws['H' + str(temp)].value
                    my_ws['G' + str(temp)] = weight * my_ws['F' + str(temp)].value  # multiplying weight * quantity
                    # need to duplicate cells.
                    r: int = int(open_q / master_p)
                    if open_q % master_p == 0:
                        for dup in range(r-1):
                            value_list.clear()
                            for i in range(1, 43):
                                value_list.append(my_ws[str(get_column_letter(i)) + str(temp)].value)
                            value_list[5] = master_p
                            value_list[6] = weight * value_list[5]
                            print(value_list)
                            my_ws.append(value_list)
                            my_ws['G' + str(my_ws.max_row)].fill = masterpack_fill  # filling OG cell range with blue color
                    else:
                        for dup in range(r):
                            value_list.clear()
                            for i in range(1, 43):
                                value_list.append(my_ws[str(get_column_letter(i)) + str(temp)].value)
                            if dup == r - 1:  # if increment is on final loop
                                value_list[5] = open_q % master_p
                            else:
                                value_list[5] = master_p
                            print(value_list)
                            value_list[6] = weight * value_list[5]
                            my_ws.append(value_list)
                            my_ws['G' + str(my_ws.max_row)].fill = masterpack_fill  # filling OG cell range with blue color

        temp -= 1
    print("...Finished")
    my_wb.save('spreadsheet_NEW.xlsx')


def get_data_coord() -> tuple:
    start = 'A1'
    _end_col = get_column_letter(my_ws.max_column)
    _end_row = my_ws.max_row
    end = (_end_col + str(_end_row))
    return start, end


def filter_OG():
    filter_red()
    filter_scac()
    filter_freight()
    filter_customer_name()
    filter_spec()
    filter_zero_weights()
    filter_char_osadx()
    duplicate_cartons()
    masterpacks()


filter_OG()
