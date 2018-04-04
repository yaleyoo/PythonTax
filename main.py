# -*- coding: utf-8 -*- #
def checking_main(wb,report):
	arr = []
	temp_D3 = 0
	main_sheet = wb.get_sheet_by_name(u'A100000主表')
	A101010 = wb.get_sheet_by_name(u'A101010一般企业')
	A101020 = wb.get_sheet_by_name(u'A101020金融企业')
	A103000 = wb.get_sheet_by_name(u'A103000事业、非营利')

	###### D3 #####
	arr.append(A101010["C3"].value)
	arr.append(A101020["C3"].value)
	arr.append(A103000["C4"].value)
	arr.append(A103000["C5"].value)
	arr.append(A103000["C6"].value)
	arr.append(A103000["C7"].value)
	arr.append(A103000["C8"].value)
	arr.append(A103000["C13"].value)
	arr.append(A103000["C14"].value)
	arr.append(A103000["C15"].value)
	arr.append(A103000["C16"].value)
	arr.append(A103000["C17"].value)

	for i in arr:
		if i==None:
			i=0
		temp_D3 = temp_D3+i

	D3 = main_sheet["D3"].value
	if D3==None:
		D3=0

	if(D3 == temp_D3):
		report.write('主表D3正确\n')
	else:
		report.write('主表D3错误：应为A101010.C3，或A101020.C3，或A103000.C4+C5+C6+C7+C8+C13+C14+C15+C16+C17\n')

	###### D4 #######
	temp_D4 = 0
	A102010 = wb.get_sheet_by_name(u'A102010一般企业')
	A102020 = wb.get_sheet_by_name(u'A102020金融企业')
	#A10300
	arr = []
	arr.append(A102010["C3"].value)
	arr.append(A102020["C3"].value)
	arr.append(A103000["C21"].value)
	arr.append(A103000["C22"].value)
	arr.append(A103000["C23"].value)
	arr.append(A103000["C24"].value)
	arr.append(A103000["C27"].value)
	arr.append(A103000["C28"].value)
	arr.append(A103000["C29"].value)

	for i in arr:
		if i==None:
			i=0
		temp_D4 = temp_D4+i

	D4 = main_sheet["D4"].value
	if D4==None:
		D4=0
	if(D4 == temp_D4):
		report.write('主表D4正确\n')
	else:
		report.write('主表D4错误，请检查\n')

	###### D6 #######
	A104000 = wb.get_sheet_by_name(u'A104000期间费用') 
	D6 = main_sheet["D6"].value
	if D6==None:
		D6=0
	temp_D6 = A104000["C29"].value

	if(D6 == temp_D6):
		report.write('主表D6正确\n')
	else:
		report.write('主表D6错误，请检查\n')

	##### D7 ####
	D7 = main_sheet["D7"].value
	if D7==None:
		D7=0
	temp_D7 = A104000["E29"].value

	if(D7 == temp_D7):
		report.write('主表D7正确\n')
	else:
		report.write('主表D7错误，请检查\n')

	### D8  ####
	D8 = main_sheet["D8"].value
	if D8==None:
		D8=0
	temp_D8 = A104000["G29"].value

	if(D8 == temp_D8):
		report.write('主表D8正确\n')
	else:
		report.write('主表D8错误，请检查\n')
	###### D12 ######
	D12 = main_sheet["D12"].value
	if D12==None:
		D12=0

	D3 = main_sheet["D3"].value
	if D3==None:
		D3 = 0
	D4 = main_sheet["D4"].value
	if D4==None:
		D4 = 0
	D5 = main_sheet["D5"].value
	if D5==None:
		D5 = 0
	D6 = main_sheet["D6"].value
	if D6==None:
		D6 = 0
	D7 = main_sheet["D7"].value
	if D7==None:
		D7 = 0
	D8 = main_sheet["D8"].value
	if D8==None:
		D8 = 0
	D9 = main_sheet["D9"].value
	if D9==None:
		D9 = 0
	D10 = main_sheet["D10"].value
	if D10==None:
		D10 = 0
	D11 = main_sheet["D11"].value
	if D11==None:
		D11 = 0

	temp_D12 = D3-D4-D5-D6-D7-D8-D9+D10+D11
	if(D12==temp_D12):
		report.write('主表D12正确\n')
	else:
		report.write('主表D12错误，请检查\n')
	############ D13 ###########
	#A101010一般企业!C18+A101020金融企业!C37+A103000事业、非营利!C11+A103000事业、非营利!C19
	D13 = main_sheet["D13"].value
	if D13==None:
		D13=0
	C18 = A101010["C18"].value
	if C18==None:
		C18=0
	C37 = A101020["C37"].value
	if C37==None:
		C37=0
	C11 = A103000["C11"].value
	if C11==None:
		C11=0
	C19 = A103000["C19"].value
	if C19==None:
		C19=0
	temp_D13 = C18+C37+C11+C19
	if(D13==temp_D13):
		report.write('主表D13正确\n')
	else:
		report.write('主表D13错误，请检查\n')
	############ D14 #######
	#A102010一般企业!C18+A102020金融企业!C35+A103000事业、非营利!C25+A103000事业、非营利!C30
	D14 = main_sheet["D14"].value
	if D14==None:
		D14=0
	C18 = A102010["C18"].value
	if C18==None:
		C18=0
	C35 = A102020["C35"].value
	if C35==None:
		C35=0
	C25 = A103000["C25"].value
	if C25==None:
		C25=0
	C30 = A103000["C30"].value
	if C30==None:
		C30=0
	temp_D14 = C18+C35+C25+C30
	if(D14==temp_D14):
		report.write('主表D14正确\n')
	else:
		report.write('主表D14错误，请检查\n')
	######## D15 ########
	#=D12+D13-D14
	D15 = main_sheet["D15"].value
	if D15==None:
		D15=0
	D12 = main_sheet["D12"].value
	if D12==None:
		D12=0
	D13 = main_sheet["D13"].value
	if D13==None:
		D13=0
	D14 = main_sheet["D14"].value
	if D14==None:
		D14=0
	temp_D15 = D12+D13-D14
	if(D15==temp_D15):
		report.write('主表D15正确\n')
	else:
		report.write('主表D15错误，请检查\n')
	###### D16 ######
	#='A108010境外所得调整'!O14-'A108010境外所得调整'!L14	
	A108010 = wb.get_sheet_by_name(u'A108010境外所得调整')
	D16 = main_sheet["D16"].value
	if D16==None:
		D16=0
	O14 = A108010["O14"].value
	if O14==None:
		O14=0
	L14 = A108010["L14"].value
	if L14==None:
		L14=0
	temp_D16 = O14-L14
	if(D16==temp_D16):
		report.write('主表D16正确\n')
	else:
		report.write('主表D16错误，请检查\n')
	###### D17 ##########
	#='A105000纳税调整'!E48
	A105000 = wb.get_sheet_by_name(u'A105000纳税调整')
	E48 = A105000["E48"].value
	if E48==None:
		E48=0
	D17 = main_sheet["D17"].value
	if D17==None:
		D17=0
	if(D17==E48):
		report.write('主表D17正确\n')
	else:
		report.write('主表D17错误，请检查\n')
	###### D18 ##########
	#='A105000纳税调整'!F48
	F48 = A105000["F48"].value
	if F48==None:
		F48=0
	D18 = main_sheet["D18"].value
	if D18==None:
		D18=0
	if(D18==F48):
		report.write('主表D18正确\n')
	else:
		report.write('主表D18错误，请检查\n')
	##### D19 ##############
	#='A107010免税减计及加计'!C33
	A107010 = wb.get_sheet_by_name(u'A107010免税减计及加计')
	C33 = A107010["C33"].value
	if C33==None:
		C33=0
	D19 = main_sheet["D19"].value
	if D19==None:
		D19=0
	if(D19==C33):
		report.write('主表D19正确\n')
	else:
		report.write('主表D19错误，请检查\n')
	#### D20 ########
	#=IF(D15-D16+D17-D18-D19>=0,0,IF('A108000境外所得税收抵免'!F14<=0,0,IF('A108000境外所得税收抵免'!F14>0,MIN(ABS('A108000境外所得税收抵免'!F14),ABS(D15-D16+D17-D18-D19)),0)))
	D20 = main_sheet["D20"].value
	if D20==None:
		D20=0
	D15 = main_sheet["D15"].value
	if D15==None:
		D15=0
	D16 = main_sheet["D16"].value
	if D16==None:
		D16=0
	D17 = main_sheet["D17"].value
	if D17==None:
		D17==0
	D18 = main_sheet["D18"].value
	if D18==None:
		D18==0
	D19 = main_sheet["D19"].value
	if D19==None:
		D19==0
	A108000 = wb.get_sheet_by_name(u'A108000境外所得税收抵免')
	F14 = A108000["F14"].value

	#如果一个if为true，之后的判断不会被执行
	if D15-D16+D17-D18-D19>=0:
		temp_D20=0
	else:
		if F14<=0:
			temp_D20=0
		else:
			if F14>0:
				#MIN(ABS('A108000境外所得税收抵免'!F14),ABS(D15-D16+D17-D18-D19)
				if abs(F14)>= abs(D15-D16+D17-D18-D19):
					temp_D20 = abs(D15-D16+D17-D18-D19)
				else:
					temp_D20 = abs(F14)
			else:
				temp_D20=0
	if(D20==temp_D20):
		report.write('主表D20正确\n')
	else:
		report.write('主表D20错误，请检查\n')
	## D21 ######
	#=D15-D16+D17-D18-D19+D20
	D21 = main_sheet["D21"].value
	temp_D21 = D15-D16+D17-D18-D19+D20
	if D21==temp_D21:
		report.write('主表D21正确\n')
	else:
		report.write('主表D21错误，请检查\n')
	### D22 ######## 																#####yaogai
	#=IF(D21<=0,0,D20 MIN(D21,'A107020所得税减免优惠'!M26))
	D22 = main_sheet["D22"].value
	if D21<=0:
		temp_D22=0
	else:
		temp_D22=D20

	### D23 #######
	#='A106000企业所得税弥补亏损'!L10
	A106000 = wb.get_sheet_by_name(u'企业所得税弥补亏损')
	L10 = A106000["L10"].value
	if L10==None:
		L10=0
	D23 = main_sheet["D23"].value
	if D23==L10:
		report.write('主表D23正确\n')
	else:
		report.write('主表D23错误，请检查\n')
	########## D24 ############
	#='A107030抵扣应纳税'!C21
	A107030 = wb.get_sheet_by_name(u'A107030抵扣应纳税')
	D24 = main_sheet["D24"].value
	if D24==None:
		D24=0
	C21 = A107030["C21"].value
	if C21==None:
		C21=0	
	if D24==C21:
		report.write('主表D24正确\n')
	else:
		report.write('主表D24错误，请检查\n')
	########## D25 ############
	=IF(D21-D22-D23-D24>=0,D21-D22-D23-D24,0)
	D25 = main_sheet["D25"].value
	