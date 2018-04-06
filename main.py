# -*- coding: utf-8 -*-
import wx 
import mainFrame

class MyFrame1(mainFrame.MyFrame1): 
   def __init__(self,parent): 
      mainFrame.MyFrame1.__init__(self,parent)  
        
app = wx.App(False) 
frame = MyFrame1(None) 
frame.Show(True) 
# 主循环
app.MainLoop()