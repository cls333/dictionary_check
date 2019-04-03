# -*- coding:utf-8 -*-


import wx
# import os
from comb import *



# import comb_v1.3 as comb


class DCFrame(wx.Frame):
    """
    数据字典稽核
    """

    def __init__(self, *args, **kw):
        # ensure the parent's __init__ is called
        super(DCFrame, self).__init__(*args, **kw)
        self.SetSize(900, 700)

        # create a panel in the frame
        pnl = wx.Panel(self)

        v_box = wx.BoxSizer(wx.VERTICAL)
        # 布局文件选择部件
        file_selector_layout = self.fileSelectorLayout(pnl)
        v_box.Add(file_selector_layout, proportion=0.5, flag=wx.EXPAND | wx.ALL)
        # 布局规则选择部件
        rule_selector_layout = self.ruleSelectorLayout(pnl)
        v_box.Add(rule_selector_layout, proportion=4, flag=wx.EXPAND | wx.ALL)
        # 布局运行日志部件
        running_log_layout = self.runningLogLayout(pnl)
        v_box.Add(running_log_layout, proportion=5, flag=wx.EXPAND | wx.ALL)
        # 布局执行按钮部件
        exec_button_layout = self.execButtonLayout(pnl)
        v_box.Add(exec_button_layout, proportion=0.5, flag=wx.EXPAND | wx.ALL)

        pnl.SetSizer(v_box)

    def companyLogoLayout(self, p):
        jpg = wx.Image('logo.jpeg', wx.BITMAP_TYPE_JPEG)
        return jpg

    def fileSelectorLayout(self, p):
        self.path = wx.TextCtrl(p, style=wx.TE_READONLY)
        selector = wx.Button(p, label="选择文件")
        selector.Bind(wx.EVT_BUTTON, self.OnFileSelector)
        sbox = wx.StaticBox(p, -1, '字典文件')
        sbsizer = wx.StaticBoxSizer(sbox)
        # proportion：相对比例
        # flag：填充的样式和方向,wx.EXPAND为完整填充，wx.ALL为填充的方向
        # border：边框
        sbsizer.Add(self.path, proportion=8, flag=wx.EXPAND | wx.ALL, border=2)  # 添加组件
        sbsizer.Add(selector, proportion=2, flag=wx.EXPAND | wx.ALL, border=2)  # 添加组件
        return sbsizer

    def ruleSelectorLayout(self, p):
        self.checkboxes = []
        sbox = wx.StaticBox(p, -1, '检核规则')
        sbsizer = wx.StaticBoxSizer(sbox, wx.VERTICAL)
        self.rules = {'r01_tab': '表级信息存在空值', 'r02_col': '字段级信息存在空值', 'r03_cod': '代码级信息存在空值',
                      'r04_tab': '表级信息表中文名存在非法字符', 'r05_col': '字段级信息字段中文名存在非法字符',
                      'r06_cod': '代码级信息代码值存在非法字符', 'r07_tab': '表级信息存在重复表', 'r08_col': '字段级信息存在重复字段',
                      'r09_cod': '代码级信息存在重复代码值', 'r10_tab': '表级信息中表的主键字段在字段级信息中不存在',
                      'r11_col': '字段级信息中字段长度为1未标记为代码字段', 'r12_col': '字段级信息标记为码值的字段只有一个码值',
                      'r13_col': '字段级信息中非代码字段存在码值', 'r14_tab': '表级信息中存在的表没有字段信息',
                      'r15_col': '字段级信息中存在的表没有表信息', 'r16_cod': '代码级信息中代码值和注释完全相同',
                      'r17_tab': '表级信息中表英文名和中文名相同', 'r18_col': '字段级信息中字段英文名和中文名相同',
                      'r19_col': '字段级信息字段中文名相同但是否代码字段标识不同'}
        for i in self.rules.keys():
            setattr(self, 'checkbox_' + i, wx.CheckBox(sbox, -1, self.rules[i]))
            eval('self.' + 'checkbox_' + i).SetValue(True)
            sbsizer.Add(eval('self.' + 'checkbox_' + i), border=2)
            self.checkboxes.append(eval('self.' + 'checkbox_' + i))
        # print(self.checkboxes)
        return sbsizer

    def runningLogLayout(self, p):
        sbox = wx.StaticBox(p, -1, '检核日志')
        sbsizer = wx.StaticBoxSizer(sbox)
        self.log = wx.TextCtrl(p, style=wx.TE_READONLY | wx.TE_DONTWRAP)
        sbsizer.Add(self.log, proportion=10, flag=wx.EXPAND | wx.ALL, border=2)
        return sbsizer

    def execButtonLayout(self, p):
        sbox = wx.StaticBox(p, -1, '')
        sbsizer = wx.StaticBoxSizer(sbox)
        self.startbutton = wx.Button(p, label='开始')
        # self.cancelbutton = wx.Button(p, label='取消')
        self.startbutton.Bind(wx.EVT_BUTTON, self.OnStart)
        # self.cancelbutton.Bind(wx.EVT_BUTTON, self.OnCancel)
        sbsizer.Add(self.startbutton, proportion=1, flag=wx.EXPAND | wx.ALL, border=2)
        # sbsizer.Add(self.cancelbutton, proportion=1, flag=wx.EXPAND | wx.ALL, border=2)

        return sbsizer

    def OnFileSelector(self, event):
        fd = wx.FileDialog(self, '选择文件', defaultDir=os.getcwd(), style=wx.FD_OPEN)
        if fd.ShowModal() == wx.ID_OK:
            self.path.SetValue(fd.GetPaths()[0])

    def OnStart(self, event):
        # self.log.SetValue(str(round(time.time() * 1000)) + '\n')
        rule_list = []
        if self.path.GetValue() == '':
            wx.MessageBox("请选择稽核文件", "错误警告", wx.OK | wx.ICON_ERROR)

        else:
            self.log.Clear()
            self.log.AppendText(str(round(time.time() * 1000)) + '\n')
            self.log.Update()
            # wx.CallAfter(self.log.Update)
            self.startbutton.Disable()
            # print('pp', path)
            for v in self.rules.keys():
                if eval('self.' + 'checkbox_' + v).GetValue():
                    rule_list.append(eval(v))
            # 这一段实现实时更新好难
            for i in running(input_file=self.path.GetValue(), func_list=rule_list):
                print(i)
                # self.log.SetValue(self.log.GetValue() + i + '\n')
                self.log.AppendText(i + '\n')
                self.log.Update()
                # wx.CallAfter(self.log.Update)

        self.startbutton.Enable()


if __name__ == '__main__':
    # When this module is run (not imported) then create the app, the
    # frame, show it, and start the event loop.
    app = wx.App()
    frm = DCFrame(None, title='数据字典检核')
    frm.Show()
    app.MainLoop()
