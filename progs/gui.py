import os
from prog1 import prog1
from prog2 import prog2
from prog3 import prog3
from prog4 import prog4
from konfs.konfs import labels, getter

import wx


class MyFrame(wx.Frame):
    def __init__(self):
        self.folder = None
        self.pathname = None
        self.archive = None
        self.default_svodki_path = r"..\1_Сводки"
        self.defaultActPath = r"..\2_Журнал_и_акты"
        no_resize = wx.DEFAULT_FRAME_STYLE & ~ (wx.RESIZE_BORDER | wx.MAXIMIZE_BOX)
        super().__init__(parent=None, title=labels[4], size=(500, 291), style=no_resize)
        self.Center()
        panel = wx.Panel(self)
        ico = wx.Icon(r'..\konfs\Шаблоны\NSHL_icon.ico', wx.BITMAP_TYPE_ICO)
        self.SetIcon(ico)
        my_sizer = wx.BoxSizer(wx.VERTICAL)

        m_staticText = wx.StaticText(panel, wx.ID_ANY, u"Выберите операцию:", wx.DefaultPosition, wx.DefaultSize,
                                     0)  # Некомпилированная версия.
        m_staticText.Wrap(-1)
        m_staticText.SetFont(
            wx.Font(15, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        m_staticText2 = wx.StaticText(panel, wx.ID_ANY, u"Дополнительно:", wx.DefaultPosition, wx.DefaultSize, 0)
        m_staticText2.Wrap(-1)
        m_staticText2.SetFont(
            wx.Font(15, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL, False, wx.EmptyString))
        my_btn3 = wx.Button(panel, label=labels[0])
        my_btn3.Bind(wx.EVT_BUTTON, self.on_press3, id=my_btn3.GetId())
        my_btn4 = wx.Button(panel, label=labels[1])
        my_btn4.Bind(wx.EVT_BUTTON, self.on_press2, id=my_btn4.GetId())
        my_btn1 = wx.Button(panel, label=labels[2])
        my_btn1.Bind(wx.EVT_BUTTON, self.on_press1, id=my_btn1.GetId())
        my_btn = wx.Button(panel, label=labels[3])
        my_btn.Bind(wx.EVT_BUTTON, self.OnSelectFile, id=my_btn.GetId())
        my_btn5 = wx.Button(panel, label=labels[5])
        my_btn5.Bind(wx.EVT_BUTTON, self.send_mail, id=my_btn5.GetId())
        m_staticline = wx.StaticLine(panel, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, wx.LI_HORIZONTAL)
        my_sizer.Add(m_staticText, 0, wx.ALL, 5)
        my_sizer.Add(my_btn1, 0, wx.ALL | wx.CENTER, 5)
        my_sizer.Add(my_btn4, 0, wx.ALL | wx.CENTER, 5)
        my_sizer.Add(my_btn3, 0, wx.ALL | wx.CENTER, 5)
        my_sizer.Add(m_staticline, 0, wx.EXPAND | wx.ALL, 5)
        my_sizer.Add(m_staticText2, 0, wx.ALL, 5)
        my_sizer.Add(my_btn5, 0, wx.ALL | wx.CENTER, 5)
        my_sizer.Add(my_btn, 0, wx.ALL | wx.CENTER, 5)
        panel.SetSizer(my_sizer)
        self.Show()

    def OnSelectFile(self, event):  # 'Выбрать файл НШЛ'
        fileDialog = wx.FileDialog(self, "Выбрать файл журнала...", wildcard="Файл Excel (*.xlsx)|*.xlsx",
                                   style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST, defaultDir=self.defaultActPath)
        if fileDialog.ShowModal() == wx.ID_CANCEL:
            return
        self.pathname = fileDialog.GetPath()
        print(self.pathname)

    def on_press1(self, event):  # 'Создать файл из сводок'
        folder = self.folder
        # print(folder)
        if folder is None:
            dlg = wx.DirDialog(None, message="Выберите расположение папки со сводками",
                               defaultPath=self.default_svodki_path)
            dlg.Center()
            if dlg.ShowModal() == wx.ID_OK:
                print('Selected files are: ', dlg.GetPath())
                dlg1 = wx.MessageDialog(self, 'Сейчас начнется создание файла, нажмите ОК и ожидайте завершения выполнения',
                                        'Внимание',
                                        wx.OK)
                dlg1.ShowModal()
                a = prog1(dlg.GetPath())
                self.pathname = a
                dlg = wx.MessageBox('Открыть созданный файл журнала?', 'Файл журнала создан',
                                    wx.YES_NO | wx.NO_DEFAULT)
                if dlg == wx.YES:
                    os.startfile(a)
                print('Файл НШЛ создан')
        else:
            a = prog1(folder)
            dlg = wx.MessageDialog(self, 'Файл НШЛ создан', 'Уведомление',
                                   wx.OK | wx.ICON_QUESTION)
            dlg.ShowModal()
            print('Файл НШЛ создан')
            os.startfile(a)
            self.pathname = a

    def on_press2(self, event):  # 'Откорректировать журнал'
        file = self.pathname
        if file is None:
            dlg = wx.MessageDialog(self, 'Файл журнала еще не выбран или не создан в данном сеансе.', 'Уведомление',
                                   wx.OK | wx.ICON_QUESTION)
            dlg.ShowModal()
            print('Не выбран файл НШЛ')
        else:
            prog2(file)
            dlg = wx.MessageBox('Открыть откорректированный файл журнала?', 'Файл журнала откорректирован',
                                wx.YES_NO | wx.NO_DEFAULT)
            if dlg == wx.YES:
                os.startfile(file)
                print('Журнал НШЛ откорректирован')

    def on_press3(self, event):  # 'Создать акта из файла'
        file = self.pathname
        # print(file)
        if file is None:
            dlg = wx.MessageDialog(self, 'Файл журнала еще не выбран или не создан в данном сеансе.', 'Уведомление',
                                   wx.OK | wx.ICON_QUESTION)
            dlg.ShowModal()
        else:
            dlg1 = wx.MessageDialog(self, 'Сейчас начнется создание актов, нажмите ОК и ожидайте завершения выполнения', 'Внимание',
                                    wx.OK)
            dlg1.ShowModal()
            folderacts = prog3(file)
            self.archive = folderacts[1]
            dlg = wx.MessageBox('Открыть папку с актами?', 'Акты НШЛ созданы',
                                wx.YES_NO | wx.NO_DEFAULT)
            if dlg == wx.YES:
                print('Акты НШЛ созданы')
                os.system(f"explorer.exe {folderacts[0]}")

    def open_xlsx(self, event):  # 'Открыть файл НШЛ'
        file = self.pathname
        # print(file)
        if file is None:
            dlg = wx.MessageDialog(self, 'Файл журнала еще не выбран или не создан в данном сеансе', 'Уведомление',
                                   wx.OK | wx.ICON_QUESTION)
            dlg.ShowModal()
        else:
            os.startfile(file)

    def send_mail(self, event):  # 'Открыть файл НШЛ'
        file = self.archive
        # print(file)
        if file is None:
            dlg = wx.MessageDialog(self, 'Файл архива еще не создан в данном сеансе', 'Уведомление',
                                   wx.OK | wx.ICON_QUESTION)
            dlg.ShowModal()
        else:
            dlg1 = wx.MessageDialog(self, f'Производится отправка сообщения на адрес {getter}, нажмите ОК и ожидайте завершения', 'Внимание',
                                    wx.OK)
            dlg1.ShowModal()
            prog4(file)
            dlg = wx.MessageBox('Открыть файл архива?', 'Сообщение отправлено',
                                wx.YES_NO | wx.NO_DEFAULT)
            if dlg == wx.YES:
                print('Открытие архива')
                os.system(f"explorer.exe {file}")


if __name__ == '__main__':
    app = wx.App()
    frame = MyFrame()
    app.MainLoop()
