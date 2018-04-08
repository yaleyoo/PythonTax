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

def check_7040(filename, rules, flag, report):
    wb = load_workbook(filename, data_only=True)
    full_sheetname = get_sheet_names_dict(rules, wb.sheetnames)
    wb = {short_name: wb[full_sheetname[short_name]]
          for short_name in full_sheetname}
    totalTax = wb["A100000"]["D25"].value
    #### 小微企业 ###
    if flag:
        report.write("--------------------小微企业A107040表审查-----------------\n")
        c3_value = wb["A107040"]["C3"].value
        if c3_value==None:
            report.write(report_line("A107040", "C3", "错误：小微企业必须填写该单元格","\n"))
        else:
            c3_value = '{:.2f}'.format(Decimal(c3_value))

            c3_expected = totalTax*0.15
            c3_expected = '{:.2f}'.format(Decimal(c3_expected))

            if c3_value == c3_expected:
                report.write(report_line("A107040", "C3", "正确！"))
            else:
                report.write(report_line("A107040", "C3", "错误：值: ", c3_value, " 期望值: ", c3_expected,"\n"))

        c25_value = wb["A107040"]["C25"].value
        if c25_value==None:
            report.write(report_line("A107040", "C25", "错误：小微企业必须填写该单元格","\n"))
        else:
            c25_value = '{:.2f}'.format(Decimal(c25_value))

            c25_expected = totalTax*0.025
            c25_expected = '{:.2f}'.format(Decimal(c25_expected))

            if c25_value == c25_expected:
                report.write(report_line("A107040", "C25", "正确！"))
            else:
                report.write(report_line("A107040", "C25", "错误：值: ", c25_value, " 期望值: ", c25_expected,"\n"))

        c37_value = wb["A107040"]["C37"].value
        if c37_value==None:
            report.write(report_line("A107040", "C37", "错误：小微企业必须填写该单元格","\n"))
        else:
            c37_value = '{:.2f}'.format(Decimal(c37_value))
            c37_expected = totalTax*0.03
            c37_expected = '{:.2f}'.format(Decimal(c37_expected))

            if c37_value == c37_expected:
                report.write(report_line("A107040", "C37", "正确！"))
            else:
                report.write(report_line("A107040", "C37", "错误：值: ", c37_value, " 期望值: ", c37_expected,"\n"))
    elif not flag:
        report.write("--------------------非小微企业A107040表审查-----------------\n")
        c25_value = wb["A107040"]["C25"].value
        if c25_value==None:
            report.write(report_line("A107040", "C25", "错误：小微企业必须填写该单元格","\n"))
        else:
            c25_value = '{:.2f}'.format(Decimal(c25_value))
            c25_expected = totalTax*0.1
            c25_expected = '{:.2f}'.format(Decimal(c25_expected))

            if c25_value == c25_expected:
                report.write(report_line("A107040", "C25", "正确！"))
            else:
                report.write(report_line("A107040", "C25", "错误：值: ", c25_value, " 期望值: ", c25_expected,"\n"))

        c37_value = wb["A107040"]["C37"].value
        if c37_value==None:
            report.write(report_line("A107040", "C37", "错误：小微企业必须填写该单元格","\n"))
        else:
            c37_value = '{:.2f}'.format(Decimal(c37_value))
            c37_expected = totalTax*0.06
            c37_expected = '{:.2f}'.format(Decimal(c37_expected))

            if c37_value == c37_expected:
                report.write(report_line("A107040", "C37", "正确！"))
            else:
                report.write(report_line("A107040", "C37", "错误：值: ", c37_value, " 期望值: ", c37_expected,"\n"))

def check_workbook(filename, rules, report):
    report.write("-------------------------------------------------------------\n")
    report.write("|-------------------税务申报表自动审查报告------------------|\n")
    report.write("-------------------------------------------------------------\n\n")
    wb = load_workbook(filename, data_only=True)

    full_sheetname = get_sheet_names_dict(rules, wb.sheetnames)

    wb = {short_name: wb[full_sheetname[short_name]]
          for short_name in full_sheetname}

    for sheet_name in sorted(rules.keys()):
        if sheet_name not in wb:
            report.write(report_line(sheet_name, None, " 子表不存在"))
            continue
        for col in sorted(rules[sheet_name].keys()):
            for row in sorted(rules[sheet_name][col].keys()):
                cell = col + str(row)
                try:
                    value = wb[sheet_name][cell].value
                    if value==None:
                        value=0
                    value = '{:.2f}'.format(Decimal(value)) 
                    except_value = eval(rules[sheet_name][col][row])
                    except_value = '{:.2f}'.format(Decimal(except_value)) 
                    if value == except_value:
                    	#正确
                        #report.write(report_line(sheet_name, cell, "正确!"))
                        None
                    else:
                        #report.write(report_line(sheet_name, cell, "错误：值: ", value, " 期望值: ", except_value,
                         #                       " 对应公式: ", rules[sheet_name][col][row],"\n"))
                         report.write(report_line(sheet_name, cell, "错误：值: ", value, " 期望值: ", except_value,"\n"))
                except TypeError as e:
                    report.write(report_line(sheet_name, cell, "TypeError ", "message: ", str(e), " eval: ",
                                            rules[sheet_name][col][row],"\n"))
                except ValueError as e:
                    report.write(report_line(sheet_name, cell, "ValueError ", "message: ", str(e), "eval: ",
                                            rules[sheet_name][col][row],"\n"))
                except ZeroDivisionError as e:
                    #report.write(report_line(sheet_name, cell, "ZeroDivisionError ", "message: ", str(e), "eval: ",
                    #                        rules[sheet_name][col][row],"\n"))
                    #TODO
                    except_value='{:.2f}'.format(Decimal("0")) 
                    if value == except_value:
                        #report.write(report_line(sheet_name, cell, "正确!"))
                        None
                    else:
                        #report.write(report_line(sheet_name, cell, "错误：值: ", value, " 期望值: ", except_value,
                         #                       " 对应公式: ", rules[sheet_name][col][row],"\n"))
                        report.write(report_line(sheet_name, cell, "错误：值: ", value, " 期望值: ", except_value,"\n"))
                         #                       " 对应公式: ", rules[sheet_name][col][row],"\n"))

                except Exception as e:
                    # print(e)
                    # print("error on ", col, row, "eval: ", rules[sheet_name][col][row])
                    report.write(report_line(sheet_name, cell, "出错了 ", "错误信息: ", str(e), "公式: ",
                                            rules[sheet_name][col][row],"\n"))


#if __name__ == '__main__':

 #   report = open(resource_path('resources/rules.txt'), 'w')
  #  rules = load_rules()
   # check_workbook(u"所得税申报表-testing.xlsx", rules, report)

    #report.close()


