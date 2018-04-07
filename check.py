# -*- coding: utf-8 -*- #

from utilities import *
import sys
import os

def load_rules(filename = 'rules.txt'):
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(".")
    filename = os.path.join(base_path, filename)

    rules_file = open(filename, "r")
    rules_str = rules_file.read()
    rules = eval(rules_str)
    return rules

def check_workbook(filename, rules, report):

    wb = load_workbook(filename, data_only=True)

    full_sheetname = get_sheet_names_dict(rules, wb.sheetnames)

    wb = {short_name: wb[full_sheetname[short_name]]
          for short_name in full_sheetname}

    for sheet_name in rules:
        if sheet_name not in wb:
            report.write(report_line(sheet_name, None, " 子表不存在"))
            continue
        for col in rules[sheet_name]:
            for row in rules[sheet_name][col]:
                cell = col + str(row)
                try:
                    value = wb[sheet_name][cell].value
                    except_value = eval(rules[sheet_name][col][row])
                    if value == except_value:
                        report.write(report_line(sheet_name, cell, "正确!"))
                    else:
                        report.write(report_line(sheet_name, cell, "错误：值: ", value, " 期望值: ", except_value,
                                                " eval: ", rules[sheet_name][col][row],"\n"))
                except TypeError as e:
                    report.write(report_line(sheet_name, cell, "TypeError ", "message: ", str(e), " eval: ",
                                            rules[sheet_name][col][row],"\n"))
                except ValueError as e:
                    report.write(report_line(sheet_name, cell, "ValueError ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row],"\n"))
                except ZeroDivisionError as e:
                    report.write(report_line(sheet_name, cell, "ZeroDivisionError ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row],"\n"))

                except Exception as e:
                    # print(e)
                    # print("error on ", col, row, "eval: ", rules[sheet_name][col][row])
                    report.write(report_line(sheet_name, cell, "Other Error ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row],"\n"))


if __name__ == '__main__':

    report = open(resource_path('resources/rules.txt'), 'w')
    rules = load_rules()
    check_workbook(u"所得税申报表-testing.xlsx", rules, report)

    report.close()


