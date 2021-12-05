# Excel Organizer
# Drew Foster, Python, 11/21/2021
import re
from typing import Any

import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.descriptors import Bool
from openpyxl.utils import get_column_letter
from openpyxl.styles import Color, PatternFill, Font, Border

#  getting info about workbook and active worksheet
my_wb = load_workbook('spreadsheet_template_.xlsx')
my_ws = my_wb.active
end_row = my_ws.max_row
end_col = my_ws.max_column
my_ws.title = 'MAIN'

#  setting up new worksheets for fedex and UPS
my_wb.create_sheet("FEDEX")
my_wb.create_sheet("UPS")
fedex_worksheet = my_wb["FEDEX"]
ups_worksheet = my_wb["UPS"]

#  initializing var to blue filler color
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


def move_over_sheet(data, sheet):
    sheet.append(data)


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
    walmart_1 = "WAL - MART - --- EDI"

    _end_row = my_ws.max_row
    temp = _end_row
    print("Starting...")
    for row in range(1, _end_row):
        customer_name = my_ws['AP' + str(temp)].value
        if customer_name == "" or customer_name is None:
            print('AP{} Customer Name: {}. Continuing'.format(temp, customer_name))
        elif customer_name == dunhams or customer_name == bj or customer_name == walmart or customer_name == scheels or customer_name == walmart_1:
            print('AP{} Customer Name: {}. Deleting'.format(temp, customer_name))
            my_ws.delete_rows(temp)
        else:
            print('A{} Customer Name: {}. Continuing'.format(temp, customer_name))
        temp -= 1
    print("...Finished")


def filter_spec():
    value_list = []
    pilot = "SHIP VIA PILOT"
    esfww = "SHIP VIA EFWW-ESTES FORWARDING WW"
    _end_row = my_ws.max_row
    temp = _end_row
    sheet: Any
    transfer: Bool
    print("Starting...")
    for row in range(1, _end_row):
        transfer = False
        value_list.clear()
        spec = my_ws['AH' + str(temp)].value
        if spec == "" or spec is None:
            print('AH{} SPEC INSTR: {}. Continuing'.format(temp, spec))
        elif spec == esfww or spec == pilot:
            print('AH{} SPEC INSTR: {}. Deleting'.format(temp, spec))
            my_ws.delete_rows(temp)
        else:
            if spec == "SHIP VIA FEDEX HOME DELV" or spec == "SHIP VIA FEDEX GND" or spec == "SHIP VIA FEDERAL EX-STANDARD GROUND" or spec == "SHIP VIA FEDEX HOME DELIVERY" or spec == "SHIP VIA FEDEX HOME DELIVERY FEDH" or spec == "SHIP VIA FEDERAL EX-GROUND":
                transfer = True
                sheet = fedex_worksheet
            elif spec == "SHIP VIA UPS GROUND":
                transfer = True
                sheet = ups_worksheet
        if transfer:
            print(f"Moving row over to sheet {sheet}")
            for i in range(1, 43):  # Copying original cell values
                value_list.append(my_ws[str(get_column_letter(i)) + str(temp)].value)
            my_ws.delete_rows(temp)
            move_over_sheet(value_list, sheet)
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
        if weight == 0 or weight == 0.0 or weight == .0 or weight is None:
            print('G{} Weight: {}. Deleting'.format(temp, weight))
            my_ws.delete_rows(temp)
        else:
            print('G{} Weight: {}. Continuing'.format(temp, weight))
            my_ws['G' + str(temp)] = int("{:.0f}".format(weight))

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


def get_data_coord() -> tuple:
    start = 'A1'
    _end_col = get_column_letter(my_ws.max_column)
    _end_row = my_ws.max_row
    end = (_end_col + str(_end_row))
    return start, end


def filter_ship_via():
    value_list = []
    fx = "FEDEX"
    ups = "UP"
    _end_row = my_ws.max_row
    temp = _end_row
    sheet: Any
    transfer: Bool
    print("Starting... (SHIP VIA)")
    for row in range(1, _end_row):
        transfer = False
        sheet = None
        ship_via_value = my_ws['X' + str(temp)].value

        if ship_via_value == "" or ship_via_value is None:
            print('A{} SCAC: {}. Continuing'.format(temp, ship_via_value))
        elif ship_via_value == fx:
            #  move to fedex sheet
            transfer = True
            sheet = fedex_worksheet
            print('A{} SHIP_VIA: {}. Transferring data'.format(temp, ship_via_value))
        elif ship_via_value[0:2] == ups:
            #  move to ups sheet
            transfer = True
            sheet = ups_worksheet
            print('A{} SHIP_VIA: {}. Transferring data'.format(temp, ship_via_value))

        if transfer:
            value_list.clear()
            for i in range(1, 43):
                value_list.append(my_ws[str(get_column_letter(i)) + str(temp)].value)
            print(value_list)
            my_ws.delete_rows(temp)
            move_over_sheet(value_list, sheet)

        temp -= 1
    print("...Finished")


def filter_scac():
    #  move fedex and UPS to separate sheets
    value_list = []
    fedex = "FXGR"
    ups = "UP"
    _end_row = my_ws.max_row
    temp = _end_row
    sheet: Any
    transfer: Bool

    print("Starting...")
    for row in range(1, _end_row):
        transfer = False
        sheet = None
        scac_value = my_ws['AL' + str(temp)].value

        if scac_value == "" or scac_value is None:
            print('A{} SCAC: {}. Continuing'.format(temp, scac_value))
        elif scac_value == fedex:
            #  move to fedex sheet
            transfer = True
            sheet = fedex_worksheet
            print('A{} SCAC: {}. Transfering data'.format(temp, scac_value))
        elif scac_value[0:2] == ups:
            #  move to ups sheet
            transfer = True
            sheet = ups_worksheet
            print('A{} SCAC: {}. Continuing'.format(temp, scac_value))
        else:
            my_ws.delete_rows(temp)
            print('A{} SCAC: {}. Deleting'.format(temp, scac_value))

        if transfer:
            value_list.clear()
            for i in range(1, 43):
                value_list.append(my_ws[str(get_column_letter(i)) + str(temp)].value)
            print(value_list)
            my_ws.delete_rows(temp)
            move_over_sheet(value_list, sheet)

        temp -= 1
    print("...Finished")


def duplicate_cartons(sheet):
    value_list = []
    _end_row = sheet.max_row
    temp = _end_row - 1
    print("Starting... (duplicating 2x)")
    for row in range(1, _end_row - 1):
        open_q = sheet['F' + str(temp)].value
        master_p = sheet['H{0}'.format(str(temp))].value
        if open_q > master_p == 1:
            print(f"Row {temp} duplicating")
            r = int(open_q / master_p) - 1
            print(r)
            sheet['F' + str(temp)] = 1
            for dup in range(r):
                value_list.clear()
                for i in range(1, 43):
                    value_list.append(sheet[str(get_column_letter(i)) + str(temp)].value)
                print(value_list)
                sheet.append(value_list)
            print(value_list)
        temp -= 1
    print("...Finished")


def masterpacks(sheet):
    _end_row = sheet.max_row
    temp = _end_row - 1
    value_list = []

    print("Starting... (Checking for MPs)")
    for _ in range(1, _end_row - 1):
        open_q = sheet['F' + str(temp)].value
        prod = sheet['E' + str(temp)].value

        master_p = sheet['H{0}'.format(str(temp))].value
        weight = sheet['G' + str(temp)].value

        if open_q is not None and open_q > 1 and master_p != 1:
            print(f"Row {temp} is a masterpack")
            sheet['G' + str(temp)].fill = masterpack_fill  # filling OG cell range with blue color

            if open_q % master_p == 0 or open_q > master_p:
                sheet['E' + str(temp)] = f"{prod}XMP"
            else:  # quantity is less than masterpack, but still > 1
                sheet['E' + str(temp)] = f"{prod}X{open_q}"

            if open_q > master_p:  # Need to duplicate cell
                # change original cell
                sheet['F' + str(temp)] = sheet['H' + str(temp)].value
                sheet['G' + str(temp)] = weight * sheet['F' + str(temp)].value  # multiplying weight * quantity
                # need to duplicate cells.
                r: int = int(open_q / master_p)

                if open_q % master_p == 0:  # Complete masterpack
                    for dup in range(r - 1):
                        value_list.clear()
                        for i in range(1, 43):  # Copying original cell values
                            value_list.append(sheet[str(get_column_letter(i)) + str(temp)].value)
                        # Updating duplicated cell values
                        value_list[5] = master_p
                        value_list[6] = weight * value_list[5]
                        value_list[4] = f"{prod}XMP"

                        print(value_list)
                        sheet.append(value_list)
                        sheet['G' + str(sheet.max_row)].fill = masterpack_fill  # filling OG cell range with blue color

                else:  # Broken masterpack
                    for dup in range(r):
                        value_list.clear()
                        for i in range(1, 43):
                            value_list.append(sheet[str(get_column_letter(i)) + str(temp)].value)
                        if dup == r - 1:  # if increment is on final loop
                            value_list[5] = open_q % master_p
                            value_list[4] = f"{prod}X{value_list[5]}"

                        else:
                            value_list[5] = master_p
                            value_list[4] = f"{prod}XMP"

                        print(value_list)
                        value_list[6] = weight * value_list[5]
                        sheet.append(value_list)
                        sheet['G' + str(
                            sheet.max_row)].fill = masterpack_fill  # filling OG cell range with blue color
        temp -= 1
    print("...Finished")


def duplicate_and_masterpack():
    sheets = [fedex_worksheet, ups_worksheet]
    for sheet in sheets:
        duplicate_cartons(sheet)
        masterpacks(sheet)


def filter_main():
    filter_red()
    filter_freight()  # TP, PI, P
    filter_zero_weights()
    filter_customer_name()
    filter_char_osadx()

    # Now transferring or deleting
    filter_scac()  # FXGR, UPS, etc.
    filter_spec()  # Special Instructions
    filter_ship_via()

    # Need to run these on non "main" spreadsheet
    duplicate_and_masterpack()
    # duplicate_cartons()
    # masterpacks()
    my_wb.save('spreadsheet_NEW.xlsx')


filter_main()
