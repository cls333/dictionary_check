# -*- coding:utf-8 -*-
import wx
from comb import *
import time
import os
from threading import Thread
from pubsub import pub


class CheckThread(Thread):
    """实现前端页面实时更新，后台开启业务处理线程，实时将处理信息反馈给前端页面"""

    def __init__(self, input_file, func_list):
        # 线程实例化时立即启动
        Thread.__init__(self)
        # 这里传入的是要检核的文件
        self.input_file = input_file
        # 这里传入的是要检核的规则，也就是检核规则函数清单
        self.func_list = func_list
        self.start()

    def run(self):
        """线程执行的代码"""
        # 要执行的检核规则数量，
        func_totle = len(self.func_list)
        # 用于控制按钮是否可用和进度条的计数器
        run_counter = 1
        # 开启检核函数，传入检核文件和检核规则函数清单
        for i in running(self.input_file, self.func_list):
            time.sleep(0.3)
            # 反馈检核函数的输出到主进程
            wx.CallAfter(pub.sendMessage, "update", msg=i, func_totle=func_totle, run_counter=run_counter)
            run_counter = run_counter + 1
            time.sleep(0.5)


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
        v_box.Add(rule_selector_layout, proportion=4.5, flag=wx.EXPAND | wx.ALL)

        # 布局运行日志部件
        running_log_layout = self.runningLogLayout(pnl)
        v_box.Add(running_log_layout, proportion=4, flag=wx.EXPAND | wx.ALL)

        # 布局进度条部件
        process_bar_layout = self.processBarLayout(pnl)
        v_box.Add(process_bar_layout, proportion=0.5, flag=wx.EXPAND | wx.ALL)

        # 布局执行按钮部件
        exec_button_layout = self.execButtonLayout(pnl)
        v_box.Add(exec_button_layout, proportion=0.5, flag=wx.EXPAND | wx.ALL)

        pnl.SetSizer(v_box)

    def companyLogoLayout(self, p):
        """加载公司logo"""
        jpg = wx.Image('logo.jpeg', wx.BITMAP_TYPE_JPEG)
        return jpg

    def fileSelectorLayout(self, p):
        """文件选择控件，包含一个只读文本框和一个调用文件选择器的按钮"""
        # 只读文本框，记录选择的文件
        self.path = wx.TextCtrl(p, style=wx.TE_READONLY)
        # 文件选择按钮，调用文件选择器
        self.selector = wx.Button(p, label="选择文件")
        self.selector.Bind(wx.EVT_BUTTON, self.OnFileSelector)

        # 布局文本框和按钮
        sbox = wx.StaticBox(p, -1, '字典文件')
        sbsizer = wx.StaticBoxSizer(sbox)
        # proportion：相对比例
        # flag：填充的样式和方向,wx.EXPAND为完整填充，wx.ALL为填充的方向
        # border：边框
        sbsizer.Add(self.path, proportion=8, flag=wx.EXPAND | wx.ALL, border=2)  # 添加组件
        sbsizer.Add(self.selector, proportion=2, flag=wx.EXPAND | wx.ALL, border=2)  # 添加组件
        return sbsizer

    def ruleSelectorLayout(self, p):
        self.checkboxes = []
        sbox = wx.StaticBox(p, -1, '检核规则')
        sbsizer = wx.StaticBoxSizer(sbox, wx.VERTICAL)
        # 检核规则编号和描述
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
            # 将检核规则标号加上前缀转换成参数
            setattr(self, 'checkbox_' + i, wx.CheckBox(sbox, -1, self.rules[i]))
            # 设置检核规则复选框为"选中"
            eval('self.' + 'checkbox_' + i).SetValue(True)
            sbsizer.Add(eval('self.' + 'checkbox_' + i), proportion=1 / len(self.rules), border=2)

            # 添加复选框变量到复选框列表，供后函数续读取
            self.checkboxes.append(eval('self.' + 'checkbox_' + i))
        # print(self.checkboxes)
        return sbsizer

    def runningLogLayout(self, p):
        sbox = wx.StaticBox(p, -1, '检核日志')
        sbsizer = wx.StaticBoxSizer(sbox)
        # 创建多行只读文本框，这里将显示业务处理线程函数返回的计算结果
        self.log = wx.TextCtrl(p, style=wx.TE_READONLY | wx.TE_MULTILINE)
        sbsizer.Add(self.log, proportion=10, flag=wx.EXPAND | wx.ALL, border=2)
        return sbsizer

    def processBarLayout(self, p):
        sbox = wx.StaticBox(p, -1, '')
        sbsizer = wx.StaticBoxSizer(sbox)
        self.progress = wx.Gauge(sbox, -1, 100)
        sbsizer.Add(self.progress, proportion=10, flag=wx.EXPAND | wx.ALL, border=2)
        return sbsizer

    def execButtonLayout(self, p):
        sbox = wx.StaticBox(p, -1, '')
        sbsizer = wx.StaticBoxSizer(sbox)
        self.startbutton = wx.Button(p, label='开始检核')
        # self.cancelbutton = wx.Button(p, label='取消')
        self.startbutton.Bind(wx.EVT_BUTTON, self.OnStart)
        pub.subscribe(self.updateDisplay, "update")
        # self.cancelbutton.Bind(wx.EVT_BUTTON, self.OnCancel)
        sbsizer.Add(self.startbutton, proportion=1, flag=wx.EXPAND | wx.ALL, border=2)
        # sbsizer.Add(self.cancelbutton, proportion=1, flag=wx.EXPAND | wx.ALL, border=2)

        return sbsizer

    def OnFileSelector(self, event):
        fd = wx.FileDialog(self, '选择文件', defaultDir=os.getcwd(), style=wx.FD_OPEN)
        if fd.ShowModal() == wx.ID_OK:
            self.path.SetValue(fd.GetPaths()[0])

    def updateDisplay(self, msg, func_totle, run_counter):
        self.progress.SetValue((run_counter / func_totle) * 100)
        # self.log.SetValue(self.log.GetValue() + msg)
        self.log.AppendText(msg)
        self.log.AppendText('\n')
        self.Update()
        # self.log.AppendText(msg)
        if func_totle == run_counter:
            # self.log.SetValue(self.log.GetValue() + '-' * 24)
            self.log.AppendText('-' * 24)
            self.log.AppendText('\n')
            self.log.SetValue(self.log.GetValue() + '检核结束')

            self.selector.Enable()
            self.startbutton.Enable()
        self.Update()

    def OnStart(self, event):
        # self.log.SetValue(str(round(time.time() * 1000)) + '\n')
        rule_list = []
        if self.path.GetValue() == '':
            wx.MessageBox("请输入检核文件", "错误警告", wx.OK | wx.ICON_ERROR)

        else:
            self.log.Clear()
            # self.log.SetValue('批次: ' + str(round(time.time() * 1000)))
            self.log.AppendText('批次: ' + str(round(time.time() * 1000)))
            self.log.AppendText('\n')
            self.log.AppendText('-' * 24)
            # self.log.SetValue(self.log.GetValue() + '-' * 24)
            self.log.AppendText('\n')

            self.log.Update()
            # wx.CallAfter(self.log.Update)
            self.selector.Disable()
            self.startbutton.Disable()
            # print('pp', path)
            for v in self.rules.keys():
                if eval('self.' + 'checkbox_' + v).GetValue():
                    rule_list.append(eval(v))
                    # print(eval(v))
            # 设置进度条
            self.progress.range = len(rule_list)
            # 设置进度条为1，表示已经开始检核工作，避免人为误解成假死
            self.progress.SetValue(1)
            # 这一段实现实时更新好难
            CheckThread(input_file=self.path.GetValue(), func_list=rule_list)

            # self.m_staticText2.SetLabel("线程开始")
            # event.GetEventObject().Disable()


if __name__ == '__main__':
    # When this module is run (not imported) then create the app, the
    # frame, show it, and start the event loop.
    app = wx.App()
    frm = DCFrame(None, title='数据字典检核')
    frm.Show()
    app.MainLoop()
