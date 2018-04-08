# -*- coding: utf-8 -*- #
import re
from openpyxl import load_workbook, Workbook


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
            result[search_obj.group(0).encode("utf-8")] = rules[name]
    return result


def chop_cell(cell):
    match_obj = re.search("([A-Z]{1,2})([0-9]+)", cell)
    return match_obj.group(1), int(match_obj.group(2))


# NOTE only handle A~ZZ
def col_add_one(col):
    if len(col) == 1:
        if col == "Z":
            return "AA"
        return chr(ord(col)+1)
    else:  # assume at most ZZ cols
        if col[1] == "Z":
            return chr(ord(col[0])+1) + "A"
        return col[0] + chr(ord(col[1])+1)


def get_sequence(cell, last_cell, sheet_name, exclude_first=True):
    this_col, this_row = chop_cell(cell)
    last_col, last_row = chop_cell(last_cell)

    result = ""
    while last_col <= this_col:
        row = last_row
        while row <= this_row:
            if exclude_first:
                exclude_first = False
            else:
                result += ',cval(wb["%s"]["%s%d"])' % (sheet_name, last_col, row)
            row += 1
        last_col = col_add_one(last_col)
    return result


def calculate_formula(current_sheetname, formula):
    cell_pattern = re.compile(r'(\'?(A\d+)[^\s]+?!)?(\$?([A-Z]{1,2})\$?([0-9]+))')
    formula = formula.lstrip("=")
    result_str = ""

    if re.search("\s", formula):
        print("WARNING: detected blank character in ", current_sheetname, ": ", formula)
    #if re.search("[\"']", formula):
    #    print("detected quotation mark in ", current_sheetname, ": ", formula)

    last_cell = None
    match_obj = re.search(cell_pattern, formula)
    while match_obj:
        sheet_name = match_obj.group(1)
        cell = match_obj.group(4) + match_obj.group(5)

        if not sheet_name:
            sheet_name = current_sheetname
        else:
            sheet_name = match_obj.group(2)

        value = 'cval(wb["%s"]["%s"])' % (sheet_name, cell)

        start, end = match_obj.span()
        if start == 1 and formula[0] == ":":
            result_str += get_sequence(cell, last_cell, sheet_name)
        else:
            result_str += formula[:start] + value

        last_cell = cell

        formula = formula[end:]
        match_obj = re.search(cell_pattern, formula)
    result_str += formula

    result_str = re.sub("(\d+)%", lambda x: "0."+x.group(1), result_str)

    result_str = re.sub("([^><])=", lambda x: x.group(1) + "==", result_str)

    return result_str.encode("utf-8")


def getall_formula(wb):
    rules = {}
    for sheet_name in wb.sheetnames:
        rules[sheet_name] = {}
        for col in wb[sheet_name].iter_cols():
            for cell in col:
                value = cell.value
                if value and isinstance(value, (str, unicode)) and value[0] == "=":
                    (rules[sheet_name].setdefault(cell.column, {}))[cell.row] = calculate_formula(trans_sheetname(sheet_name), value)
    return rules


wb = load_workbook(u"申报表软件要求original.xlsx", data_only=False)

rules = getall_formula(wb)
rules = trans_sheetnames(rules)

rules_file = open("rules.txt", "w+")
rules_str = str(rules)
rules_file.write(rules_str)
rules_file.close()

print(rules_str)