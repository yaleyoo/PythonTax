
from openpyxl import load_workbook, Workbook

import json
import re

def AND(*data):
    for item in data:
        if not item:
            return False
    return True

def ABS(value):
    return abs(value)

def ROUND(value, digit = 0):
    return round(value, digit)

def IF(predicate, value1, value2):
    if predicate:
        return value1
    return value2

def SUM(*data):
    return sum(data)

def MAX(*data):
    return max(data)

def MIN(*data):
    return min(data)

def cval(cell):
    if cell.value is None:
        return 0
    if cell.data_type == "b" or cell.data_type == "e":
        return None
    if cell.value == "-" or cell.value == "*":
        return 0
    #if (type(cell.value) is str) and (not cell.value.isdigit()) and cell.value != "-" and cell.value != "*":
    #    return 0
    return cell.value

def report_line(sheet_name, cell, *messages):
    report = ""
    report += sheet_name + ": " + cell + ": "
    for message in messages:
        if type(message) is not str:
            message = str(message)
        report += message
    report += "\n"
    return report


def chop_sheetname(full_name):
    search_obj = re.search("(A\d+)([^\s]+)", full_name)
    if search_obj:
        return search_obj.group(0), search_obj.group(1)
    return None, full_name

def chop_sheetnames(full_names):
    sheet_id = {}
    short_name = {}
    for full_name in full_names:
        sheet_id[full_name], short_name[full_name] = chop_sheetname(full_name)
    return sheet_id, short_name


def trans_sheetname(sheetname):
    search_obj = re.search("A\d+", sheetname)
    if search_obj:
        return search_obj.group(0)
    return sheetname

def trans_sheetnames(rules):
    result = {}
    for name in rules:
        search_obj = re.search("A\d+", name)
        if search_obj:
            result[search_obj.group(0)] = rules[name]
    return result

def get_sheet_names_dict(short_names, full_names):
    dict = {}
    for short_name in short_names:
        for full_name in full_names:
            if full_name.find(short_name) >= 0:
                dict[short_name] = full_name
    return dict



#can have bug
def get_sequence(cell, last_cell, sheet_name):
    this_row = cell[1:]
    last_row = last_cell[1:]
    while not this_row.isdigit():
        this_row = cell[1:]
    while not last_row.isdigit():
        last_row = last_cell[1:]
    this_row = int(this_row)
    last_row = int(last_row)
    result = ""
    last_row += 1
    while last_row <= this_row:
        result += ',cval(wb["%s"]["%s%d"])' % (sheet_name, cell[0], last_row)
        last_row += 1
    return result


def calculate_formula(wb, current_sheetname, formula):
    cell_pattern = re.compile(r'(\'?A[\d+][^\s]+?!)?([A-Z]{1,2}[0-9]+)')
    formula = formula.lstrip("=")
    result_str = ""

    last_cell = 0
    match_obj = re.search(cell_pattern, formula)
    while match_obj:
        sheet_name = match_obj.group(1)
        cell = match_obj.group(2)
        if not sheet_name:
            sheet_name = current_sheetname
        else:
            sheet_name = sheet_name[:-1]
            sheet_name = re.sub("\'","",sheet_name)
            sheet_name = trans_sheetname(sheet_name)

        value = 'cval(wb["%s"]["%s"])' % (sheet_name, cell)

        start, end = match_obj.span()
        if start == 1 and formula[0] == ":":
            result_str += get_sequence(cell, last_cell, sheet_name)
        else:
            result_str += formula[:start] + value
        formula = formula[end:]

        last_cell = cell

        match_obj = re.search(cell_pattern, formula)
    result_str += formula

    result_str = re.sub("(\d+)%", lambda x: "0."+x.group(1), result_str)

    result_str = re.sub("([^><])=", lambda x: x.group(1) + "==", result_str)

    result_str = re.sub("\$([A-Z]+)\$(\d+)", lambda x: 'cval(wb["%s"]["%s"])' % (current_sheetname, x.group(1)+x.group(2)), result_str)

    return result_str


def getall_formula(wb):
    rules = {}
    for sheet_name in wb.sheetnames:
        rules[sheet_name] = {}
        for col in wb[sheet_name].iter_cols():
            for cell in col:
                value = cell.value
                if value and isinstance(value, str) and value[0] == "=":
                    (rules[sheet_name].setdefault(cell.column, {}))[cell.row] = calculate_formula(wb, trans_sheetname(sheet_name), value)
    return rules






