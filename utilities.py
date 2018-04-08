from openpyxl import load_workbook, Workbook
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
    if cell.value == "*":
        return 0
    #if (type(cell.value) is str) and (not cell.value.isdigit()) and cell.value != "-" and cell.value != "*":
    #    return 0
    return cell.value


def report_line(sheet_name, cell, *messages):
    report = ""
    report += sheet_name + ": " + cell + ": "
    for message in messages:
        if not isinstance(message, str):
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


def get_sheet_names_dict(short_names, full_names):
    result = {}
    for short_name in short_names:
        for full_name in full_names:
            if full_name.find(short_name) >= 0:
                result[short_name] = full_name
    return result







