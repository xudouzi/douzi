
#encoding=utf-8
"""
import wx
class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,parent,id,title = "输入框",size=(400,300))
        #创建面板
        panel = wx.Panel(self)
        #创建输入框和文本
        self.title = wx.StaticText(panel,lable = "请输入excel表格和sheet的名字")
        self.lable_excel_1 = wx.StaticText(panel,label = "表格一的名字：",pos = (50,50))
        self.test_excel_1 = wx.TextCtrl(panel,pos=(100,50),size=(235,25),style=wx.TE_LEFT)
        self.lable_excel_2 = wx.StaticText(panel, label="表格二的名字：", pos=(50, 50))
        self.test_excel_2 = wx.TextCtrl(panel, pos=(100, 50), size=(235, 25), style=wx.TE_LEFT)
        self.lable_sheet_1 = wx.StaticText(panel, label="sheet一的名字：", pos=(50, 50))
        self.test_sheet_1 = wx.TextCtrl(panel, pos=(100, 50), size=(235, 25), style=wx.TE_LEFT)
        self.lable_sheet_2 = wx.StaticText(panel, label="sheet一的名字：", pos=(50, 50))
        self.test_sheet_2 = wx.TextCtrl(panel, pos=(100, 50), size=(235, 25), style=wx.TE_LEFT)
        #确认按钮
        self.bt_confirm = wx.Button(panel,label='确定',pos=(205,130))
        self.bt_cancel = wx=

if __name__ =="__main__":
    app = wx.App()
    frame = MyFrame(parent=None,id = -1)
    frame.Show()
    app.MianLoop()
"""