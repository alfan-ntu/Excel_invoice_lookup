#
# File: im_gui.py
# Brief: Entry of invoice matching GUI
# Author: alfan-ntu
# Ver.: v. 1.0a
# Date: 2021/3/30
# ToDo's: (2021/04/03)
#   1) 檢查 110.03月電子發票明細.xls 交易幣別 (京傳企業) 交易比對錯誤
#   2) Add progress bar in GUI/multi-threading display and matching process
#   3) Display match summary
#   4) Assign match results External_Sales_GUI_yyyymm.xlsx
#   5) Apply filter to the result Excel file; and freeze the top row of the result Excel file
#

from tkinter import *
from tkinter import ttk
import tkinter.filedialog as fdlg
from im_gui_constructor import SelectorPanel, OperationPanel


# Create a new FileSelDlgDemo class based on the class ttk.Frame
class FileSelDlgDemo(ttk.Frame):
    def __init__(self, isapp=True, name='fileseldlgdemo'):
        ttk.Frame.__init__(self, name=name)
        self.pack(expand=Y, fill=BOTH)
        self.app_name ="雅博會計師事務所 統一發票對帳工具"
        self.master.title(self.app_name)
        self.isapp = isapp
        self._create_gui_widgets()

    #  Overall GUI includes
    #   1) selection panel for selecting operating files, operation date duration
    #   2) operation panel for major operation buttons and log area
    def _create_gui_widgets(self):
        if self.isapp:
            self.sel_pnl_name = "統一發票與總帳比對工具"
            self.sel_pnl = SelectorPanel(self, [self.sel_pnl_name])
            self.op_pnl = OperationPanel(self)
        # self._create_demo_panel()

    def _create_demo_panel(self):
        demoPanel = Frame(self)
        demoPanel.pack(side=TOP, fill=BOTH, expand=Y)

        for item in ('open', 'save'):
            frame = ttk.Frame(demoPanel)
            lbl = ttk.Label(frame, width=20,
                            text='Select a file to {} '.format(item))
            ent = ttk.Entry(frame, width=25)
            btn = ttk.Button(frame, text='Browse...',
                             command=lambda i=item, e=ent: self._file_dialog(i, e))
            lbl.pack(side=LEFT)
            ent.pack(side=LEFT, expand=Y, fill=X)
            btn.pack(side=LEFT, padx=5)
            frame.pack(fill=X, padx='1c', pady=3)

    def _file_dialog(self, type, ent):
        # triggered when the user clicks a 'Browse' button
        fn = None
        opts = {'initialfile': ent.get(),
                'filetypes': (('Python files', '.py'),
                              ('PNG', '.png'),
                              ('Text files', '.txt'),
                              ('All files', '.*'),)}

        if type == 'open':
            opts['title'] = 'Select a file to open...'
            fn = fdlg.askopenfilename(**opts)
        else:
            # this should only return a filename; however,
            # under windows, selecting a file and hitting
            # 'Save' gives a warning about replacing an
            # existing file; although selecting 'Yes' does
            # not actually cause a 'Save'; the filename
            # is simply returned
            opts['title'] = 'Select a file to save...'
            fn = fdlg.asksaveasfilename(**opts)

        if fn:
            ent.delete(0, END)
            ent.insert(END, fn)


if __name__ == '__main__':
    FileSelDlgDemo().mainloop()
