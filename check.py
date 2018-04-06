# -*- coding: utf-8 -*- #

from utilities import *

def load_rules(filename = "rules.txt"):
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
            report.write(report_line(sheet_name, None, " sheet is not in the workbook"))
            continue
        for col in rules[sheet_name]:
            for row in rules[sheet_name][col]:
                cell = col + str(row)
                try:
                    value = wb[sheet_name][cell].value
                    except_value = eval(rules[sheet_name][col][row])
                    if value == except_value:
                        report.write(report_line(sheet_name, cell, "SUCCESS: value: ", value, " except_value: ",
                                                except_value, " eval: ", rules[sheet_name][col][row]))
                    else:
                        report.write(report_line(sheet_name, cell, "值: ", value, " 期望值: ", except_value,
                                                " eval: ", rules[sheet_name][col][row]))
                except TypeError as e:
                    report.write(report_line(sheet_name, cell, "TypeError ", "message: ", str(e), " eval: ",
                                            rules[sheet_name][col][row]))
                except ValueError as e:
                    report.write(report_line(sheet_name, cell, "ValueError ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row]))
                except ZeroDivisionError as e:
                    report.write(report_line(sheet_name, cell, "ZeroDivisionError ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row]))

                except Exception as e:
                    # print(e)
                    # print("error on ", col, row, "eval: ", rules[sheet_name][col][row])
                    report.write(report_line(sheet_name, cell, "Other Error ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row]))


if __name__ == '__main__':

    report = open("report.txt", 'w')
    rules = load_rules()
    check_workbook(u"所得税申报表-testing.xlsx", rules, report)

    report.close()


