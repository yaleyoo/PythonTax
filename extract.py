# -*- coding: utf-8 -*- #
from utilities import *


wb = load_workbook(u"申报表软件要求original.xlsx", data_only=False)
sheet_names = wb.sheetnames

rules = getall_formula(wb)
rules = trans_sheetnames(rules)

rules_file = open("rules.txt", "w+")
rules_str = str(rules)
rules_file.write(rules_str)
rules_file.close()

print(rules_str)