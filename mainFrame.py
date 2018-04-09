# -*- coding: utf-8 -*- 

import wx
import wx.xrc
from check import *
import sys
import os

###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1 ( wx.Frame ):
	
	def __init__( self, parent ):
		global path
		path = None
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"申报表自动审查辅助工具", pos = wx.DefaultPosition, size = wx.Size( 515,338 ), style = wx.DEFAULT_FRAME_STYLE|wx.TRANSPARENT_WINDOW )
		self.SetSizeHintsSz(wx.Size(515,338), wx.Size(515,338))
		if hasattr(sys, "_MEIPASS"):
	 		base_path = sys._MEIPASS
		else:
			base_path = os.path.abspath(".")
		bgfile = os.path.join(base_path, 'bg.jpg')
	 	self.bitmap = wx.StaticBitmap(self, -1, wx.Image(bgfile, wx.BITMAP_TYPE_ANY).ConvertToBitmap(), (0, 0)) 
		self.m_menubar1 = wx.MenuBar( 0 )
		self.m_menu1 = wx.Menu()
		self.m_menuItem1 = wx.MenuItem( self.m_menu1, wx.ID_ANY, u"打开文件", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu1.AppendItem( self.m_menuItem1 )
		
		self.m_menubar1.Append( self.m_menu1, u"打开" )

		######### 绑定响应函数
		self.Bind(wx.EVT_MENU,self.__OpenSingleFile,self.m_menuItem1) 
		
		self.m_menu2 = wx.Menu()
		self.m_menuItem2 = wx.MenuItem( self.m_menu2, wx.ID_ANY, u"使用说明", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu2.AppendItem( self.m_menuItem2 )
		######### 绑定响应函数
		self.Bind(wx.EVT_MENU,self.openManual,self.m_menuItem2) 
		
		
		self.m_menuItem3 = wx.MenuItem( self.m_menu2, wx.ID_ANY, u"版权说明", wx.EmptyString, wx.ITEM_NORMAL )
		self.m_menu2.AppendItem( self.m_menuItem3 )
		
		self.m_menubar1.Append( self.m_menu2, u"帮助" ) 
		
		self.SetMenuBar( self.m_menubar1 )
		######### 绑定响应函数
		self.Bind(wx.EVT_MENU,self.openAuthor,self.m_menuItem3) 
		
		bSizer2 = wx.BoxSizer( wx.VERTICAL )

		
		
		self.m_textCtrl1 = wx.TextCtrl( self, wx.ID_ANY, u"文件路径", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_textCtrl1.Enable( False )
		
		bSizer2.Add( self.m_textCtrl1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL|wx.EXPAND, 5 )
		
		gbSizer2 = wx.GridBagSizer( 0, 0 )
		gbSizer2.SetFlexibleDirection( wx.BOTH )
		gbSizer2.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.m_button7 = wx.Button( self.bitmap, wx.ID_ANY, u"选择文件", wx.DefaultPosition, wx.DefaultSize, 0|wx.ALWAYS_SHOW_SB )
		gbSizer2.Add( self.m_button7, wx.GBPosition( 0, 1 ), wx.GBSpan( 1, 1 ), wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		## 绑定选择文件按钮监听事件
		self.Bind(wx.EVT_BUTTON,self.__OpenSingleFile,self.m_button7)
		
		self.m_button8 = wx.Button( self.bitmap, wx.ID_ANY, u"开始审查", wx.DefaultPosition, wx.DefaultSize, 0|wx.ALWAYS_SHOW_SB )
		gbSizer2.Add( self.m_button8, wx.GBPosition( 0, 2 ), wx.GBSpan( 1, 1 ), wx.ALL|wx.ALIGN_RIGHT, 5 )
		## 绑定选择文件按钮监听事件
		self.Bind(wx.EVT_BUTTON,self.__startChecking ,self.m_button8)
		
		bSizer2.Add( gbSizer2, 0, wx.ALIGN_CENTER_HORIZONTAL, 5 )

		bSizer5 = wx.BoxSizer( wx.VERTICAL )

		gbSizer21 = wx.GridBagSizer( 0, 0 )
		gbSizer21.SetFlexibleDirection( wx.BOTH )
		gbSizer21.SetNonFlexibleGrowMode( wx.FLEX_GROWMODE_SPECIFIED )
		
		self.m_radioBtn1 = wx.RadioButton( self.bitmap, wx.ID_ANY, u"小微企业", wx.DefaultPosition, wx.DefaultSize, 0 )
		gbSizer21.Add( self.m_radioBtn1, wx.GBPosition( 0, 0 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
		
		self.m_radioBtn2 = wx.RadioButton( self.bitmap, wx.ID_ANY, u"非小微企业", wx.DefaultPosition, wx.DefaultSize, 0 )
		gbSizer21.Add( self.m_radioBtn2, wx.GBPosition( 0, 1 ), wx.GBSpan( 1, 1 ), wx.ALL, 5 )
		
		bSizer5.Add( gbSizer21, 1, wx.EXPAND, 5 )

		self.m_staticText61 = TPStaticText( self.bitmap, wx.ID_ANY, u"相关政策链接：财税〔2017〕43号")
		bSizer5.Add( self.m_staticText61, 0, wx.ALL, 5 )

		gbSizer2.Add( bSizer5, wx.GBPosition( 0, 3 ), wx.GBSpan( 1, 1 ), wx.EXPAND, 5 )
		

		self.m_staticline2 = wx.StaticLine( self.bitmap, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL )
		bSizer2.Add( self.m_staticline2, 0, wx.EXPAND |wx.ALL, 5 )
		
		self.m_staticText6 = TPStaticText( self, wx.ID_ANY, u"本软件仅支持Window平台。\n1. 按照《中华人民共和国企业所得税年度纳税申报表（A类，2017年版）》所示规范填写\n   《2017年版企业所得税申报表》。\n2.《2017年版企业所得税申报表》的格式应为xlsx，如为xls格式，请使用2007或更新版本\n    Excel将表格转换为xlsx格式。\n3. 点击“选择文件”按钮，选择填写完毕的《2017年版企业所得税申报表》。\n4. 点击“开始审查”按钮。\n5. 报告将生成在该软件所处目录下。请对照报告修改申报表。\n(注意：报告中第二个参数为错误单元格在Excel中的坐标，而不是行次)")
		#self.m_staticText6.Wrap( -1 )
		bSizer2.Add( self.m_staticText6, 1, wx.ALL, 5 )
	
		#self.m_staticText5 = TPStaticText(self, wx.ID_ANY, u"                                       拉萨经济技术开发区国家税务局")
		#bSizer2.Add( self.m_staticText5, 0, wx.ALIGN_BOTTOM|wx.ALL|wx.EXPAND, 28 )
		
		
		self.SetSizer( bSizer2 )
		self.Layout()
		
		self.Centre( wx.BOTH )


	###############
	####   打开文件
	##############
	def __OpenSingleFile(self, event):
		global path
		filesFilter = "xlsx (*.xlsx)|*.xlsx"
		fileDialog = wx.FileDialog(self, message ="选择文件", wildcard = filesFilter, style = wx.FD_OPEN)
		dialogResult = fileDialog.ShowModal()
		if dialogResult !=  wx.ID_OK:
			return
		path = fileDialog.GetPath()
		self.m_textCtrl1.SetLabel(path)

	####   打开使用说明子窗口
	def openManual(self, event):
		win = manual(self)
		win.Show(True)

	####   打开使用版权子窗口
	def openAuthor(self, event):
		win = author(self)
		win.Show(True)

	######## 检查方法 入口
	def __startChecking(self,evert):
		flag = self.m_radioBtn1.GetValue()
		no_flag = self.m_radioBtn2.GetValue()

		if not flag|no_flag:
			wx.MessageBox("请先选择是否小微企业", "错误" ,wx.OK | wx.ICON_INFORMATION) 

		if path!=None:
			report = open("report.txt", 'w')
			rules = load_rules()
			check_workbook(path,rules,report, flag)

			report.close()


			win = done(self)
			win.Show(True)
		else:
			wx.MessageBox("输入文件不能为空，请选择输入文件", "错误" ,wx.OK | wx.ICON_INFORMATION) 

	def __del__( self ):
		pass
#########################
#### 透明背景
######################
class TPStaticText(wx.StaticText):  
    """ transparent StaticText """  
    def __init__(self,parent,id,label='',  
                 pos=wx.DefaultPosition,  
                 size=wx.DefaultSize,  
                 style=wx.ALIGN_CENTRE|wx.TRANSPARENT_WINDOW,  
                 name = 'TPStaticText'):  
        style |= wx.CLIP_CHILDREN | wx.TRANSPARENT_WINDOW   
        wx.StaticText.__init__(self,parent,id,label,pos,size,style = style)  
        self.SetBackgroundStyle(wx.BG_STYLE_CUSTOM)  
        self.Bind(wx.EVT_PAINT,self.OnPaint)  
          
    def OnPaint(self,event):  
        event.Skip()  
        dc = wx.GCDC(wx.PaintDC(self) )  
        dc.SetFont(self.GetFont())                  
        dc.DrawText(self.GetLabel(), 0, 0) 
###########################################################################
## Class done
###########################################################################

class done ( wx.Dialog ):
	
	def __init__( self, parent ):
		wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = u"审查完成", pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_DIALOG_STYLE )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		
		bSizer4 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_staticText4 = wx.StaticText( self, wx.ID_ANY, u"审查报告已生成在当前文件夹！", wx.DefaultPosition, wx.DefaultSize, 0 )
		self.m_staticText4.Wrap( -1 )
		bSizer4.Add( self.m_staticText4, 1, wx.ALL|wx.EXPAND, 5 )
		
		self.m_button3 = wx.Button( self, wx.ID_ANY, u"完成", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer4.Add( self.m_button3, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
		self.Bind(wx.EVT_BUTTON,self.__done__ ,self.m_button3)
		
		self.SetSizer( bSizer4 )
		self.Layout()
		bSizer4.Fit( self )
		
		self.Centre( wx.BOTH )
	
	def __done__(self, event):
		self.Close(True)

	def __del__( self ):
		pass
	

###########################################################################
## Class author
###########################################################################

class author ( wx.Dialog ):
	
	def __init__( self, parent ):
		wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = u"版权说明", pos = wx.DefaultPosition, size = wx.Size(300,150), style = wx.DEFAULT_DIALOG_STYLE )
		
		self.SetSizeHintsSz( wx.Size(300,150), wx.Size(300,150) )
		
		bSizer7 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_staticText8 = wx.StaticText( self, wx.ID_ANY, u"本软件是一个为《2017年版企业所得税申报表》提供自动审查服务的自动审查工具，版本号：v1.0\n\n版权所有：西藏经济技术开发区国家税务局\n\n开发者：郭元昱，蒋若波，朱慧媛\n", wx.DefaultPosition, wx.Size( 300,150 ), 0 )
		self.m_staticText8.Wrap( -1 )
		bSizer7.Add( self.m_staticText8, 1, wx.ALL|wx.EXPAND, 5 )
		
		
		self.SetSizer( bSizer7 )
		self.Layout()
		
		self.Centre( wx.BOTH )
	
	def __del__( self ):
		pass
	

###########################################################################
## Class manual
###########################################################################

class manual ( wx.Dialog ):
	
	def __init__( self, parent ):
		wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = u"使用说明", pos = wx.DefaultPosition, size = wx.Size( 300,200 ), style = wx.DEFAULT_DIALOG_STYLE )
		
		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )
		
		bSizer6 = wx.BoxSizer( wx.VERTICAL )
		
		self.m_textCtrl3 = wx.TextCtrl( self, wx.ID_ANY, u"使用说明：\n本软件仅支持Window平台。\n1. 按照《中华人民共和国企业所得税年度纳税申报表（A类，2017年版）》所示规范填写《2017年版企业所得税申报表》。\n2.《2017年版企业所得税申报表》的格式应为xlsx，如为xls格式，请使用2007或更新版本Excel将表格转换为xlsx格式。\n3. 点击“选择文件”按钮，选择填写完毕的《2017年版企业所得税申报表》。\n4. 点击“开始审查”按钮。\n5. 报告将生成在该软件所处目录下。请对照报告修改申报表(注意：报告中第二个参数为错误单元格在Excel中的坐标，而不是行次)。\n----------------------------------------------------\n免责声明：\n鉴于本软件使用非人工检索/分析方式，无法确定您输入的数据是否合法，所以对审查结果不承担责任。\n----------------------------------------------------\n解释权：\n本软件之说明以及其修改权、更新权及最终解释权均属拉萨经济技术开发区国家税务局所有。", wx.DefaultPosition, wx.DefaultSize, 0|wx.VSCROLL|wx.TE_READONLY|wx.TE_MULTILINE )
		bSizer6.Add( self.m_textCtrl3, 1, wx.ALL|wx.EXPAND, 5 )

		self.SetSizer( bSizer6 )
		self.Layout()
		
		self.Centre( wx.BOTH )

	
	def __del__( self ):
		pass
	

