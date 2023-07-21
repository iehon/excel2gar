import pandas as pd
import wx
import wx.xrc
import os

class MyApp(wx.App):
    def OnInit(self):
        self.frame = MyFrame(None)
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True

class MyFrame(wx.Frame):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=wx.EmptyString, pos=wx.DefaultPosition,
                          size=wx.Size(500, 300), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)
        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.panel.SetSizer(self.sizer)

        self.openFileButton = wx.Button(self.panel, label='Open File')
        self.openFileButton.Bind(wx.EVT_BUTTON, self.OnOpenFile)
        
        self.copyButton = wx.Button(self.panel, label='Copy')
        self.copyButton.Bind(wx.EVT_BUTTON, self.OnCopy)

        self.textCtrl = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE|wx.TE_READONLY)
        self.sizer.Add(self.openFileButton, 0, wx.ALL, 5)
        self.sizer.Add(self.copyButton, 0, wx.ALL, 5)
        self.sizer.Add(self.textCtrl, 1, wx.EXPAND|wx.ALL, 5)

    def OnOpenFile(self, event):
        with wx.FileDialog(self, "Open Excel file", wildcard="Excel files (*.xls;*.xlsx)|*.xls;*.xlsx") as fileDialog:
            if fileDialog.ShowModal() == wx.ID_CANCEL:
                return
            pathname = fileDialog.GetPath()
            self.processExcelFile(pathname)

    def processExcelFile(self, pathname):
        df = pd.read_excel(pathname)
        df['Разница'] = df.iloc[:, 5] - df.iloc[:, 4]

        words_to_add = df[df['Разница'] > 0]
        words_to_add = words_to_add[['Слова', 'Разница']].sort_values('Разница', ascending=False)

        words_not_to_use = df[df['Разница'] < -2]
        words_not_to_use = words_not_to_use[['Слова', 'Разница']]

        words_to_add_text = "Используй слова в количестве:\n" + '\n'.join(words_to_add['Слова'] + '\t' + words_to_add['Разница'].astype(str))
        words_not_to_use_text = "Не используй эти слова:\n" + '\n'.join(words_not_to_use['Слова'])

        self.textCtrl.SetValue(words_to_add_text + '\n\n' + words_not_to_use_text)
    
    def OnCopy(self, event):
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(self.textCtrl.GetValue()))
            wx.TheClipboard.Close()

if __name__ == '__main__':
    app = MyApp(False)
    app.MainLoop()
