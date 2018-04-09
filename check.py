# -*- coding: utf-8 -*- #

from utilities import *
import sys
import os
from decimal import Decimal


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


def check_cell(wb, value, rule, must_fill=False, not_fill_message="必须填写此单元格"):
    try:
        if value is None:
            if must_fill:
                return not_fill_message
            value = 0
        value = '{:.2f}'.format(Decimal(value))
        if isinstance(rule, (int, float)):
            expect_value = rule
        else:
            expect_value = eval(rule)
        expect_value = '{:.2f}'.format(Decimal(expect_value))
        if value == expect_value:  # 正确
            return None
        else:
            return "错误：值: " + value + " 期望值: " + expect_value
    except TypeError as e:
        return "TypeError " + "message: " + str(e) + " eval: "+rule
    except ValueError as e:
        return "ValueError " + "message: " + str(e) + " eval: " + rule
    except ZeroDivisionError as e:
        #return "ZeroDivisionError " + "message: " + str(e) + " eval: " + rule
        # TODO
        expect_value = '{:.2f}'.format(Decimal("0"))
        if value == expect_value:
            return None
        else:
            return "错误：值: " + value + " 期望值: " + expect_value
    except Exception as e:
        return "出错了 " + "错误信息: " + str(e) + " 公式: " + rule


def check_7040(wb, flag, report):
    totalTax = wb["A100000"]["D25"].value

    # 待确认行为
    if totalTax is None:
        totalTax = 0

    #### 小微企业 ###
    if flag:
        report.write("--------------------小微企业A107040表审查-----------------\n")
        c3_value = wb["A107040"]["C3"].value
        message = check_cell(wb, c3_value, totalTax*0.15, True, "错误：小微企业必须填写该单元格")
        if message:
            report.write(report_line("A107040", "C3", message))

        c25_value = wb["A107040"]["C25"].value
        message = check_cell(wb, c25_value, totalTax*0.025, True, "错误：小微企业必须填写该单元格")
        if message:
            report.write(report_line("A107040", "C25", message))

        c37_value = wb["A107040"]["C37"].value
        message = check_cell(wb, c37_value, totalTax*0.03, True, "错误：小微企业必须填写该单元格")
        if message:
            report.write(report_line("A107040", "C37", message))
    else:
        report.write("--------------------非小微企业A107040表审查-----------------\n")
        c25_value = wb["A107040"]["C25"].value
        message = check_cell(wb, c25_value, totalTax*0.1, True, "错误：非小微企业必须填写该单元格")
        if message:
            report.write(report_line("A107040", "C25", message))

        c37_value = wb["A107040"]["C37"].value
        message = check_cell(wb, c37_value, totalTax*0.06, True, "错误：非小微企业必须填写该单元格")
        if message:
            report.write(report_line("A107040", "C37", message))

def check_general(wb, rules, report):
    for sheet_name in sorted(rules.keys()):
        if sheet_name not in wb:
            report.write(report_line(sheet_name, None, " 子表不存在"))
            continue
        for col in sorted(rules[sheet_name].keys()):
            for row in sorted(rules[sheet_name][col].keys()):
                cell = col + str(row)
                value = wb[sheet_name][cell].value
                rule = rules[sheet_name][col][row]

                message = check_cell(wb, value, rule)
                if message:
                    if sheet_name == "A107030":
                        report.write(report_line(sheet_name, cell, message + " (非享受创业投资企业抵扣应纳税所得额优惠（含结转）纳税人可忽略本条错误提醒)\n"))
                    else:
                        report.write(report_line(sheet_name, cell, message))


def check_workbook(filename, rules, report, small_business):
    report.write("-------------------------------------------------------------\n")
    report.write("|-------------------税务申报表自动审查报告------------------|\n")
    report.write("-------------------------------------------------------------\n\n")
    wb = load_workbook(filename, data_only=True)

    full_sheetname = get_sheet_names_dict(rules, wb.sheetnames)

    wb = {short_name: wb[full_sheetname[short_name]]
          for short_name in full_sheetname}

    check_general(wb, rules, report)
    check_7040(wb, small_business, report)



if __name__ == '__main__':

    path = u"非小微企业测试.xlsx"

    report = open('report.txt', 'w')
    rules = load_rules()
    check_workbook(path, rules, report, True)

    report.close()




