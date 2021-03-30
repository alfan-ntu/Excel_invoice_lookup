# File: demopanels.py
# References:
#    http://hg.python.org/cpython/file/4e32c450f438/Lib/tkinter/simpledialog.py
#    http://docs.python.org/py3k/library/inspect.html#module-inspect
#
# Icons sourced from:
#    http://findicons.com/icon/69404/deletered?width=16#
#    http://findicons.com/icon/93110/old_edit_find?width=16#
#
# This file is imported by the Tkinter Demos
#
# Author: Al Fan@yapro
# Date: 2021/3/30
# ToDo's :
#       1) adjust log_text to vertically scroll automatically when a new message is inserted
#

from tkinter import *
from tkinter import ttk
from tkinter.simpledialog import Dialog
import tkinter.font as font
import tkinter.filedialog as fdlg
from PIL import Image, ImageTk
from tkcalendar import Calendar, DateEntry
from datetime import datetime
import constant


# SelectorPanel class creates the upper half of the GUI, which includes
# 1. label/entry/button for selecting the original invoice record
# 2. label/entry/button for selecting the original general ledger record
# 3. label/date selector of the starting date; label/date selector of the end date
class SelectorPanel(ttk.Frame):
    # SelectorPanel outputs the following operation parameters
    # invoice_ent: the full path to the invoice (統一發票開票紀錄檔)
    # gl_ent: the full path to the general ledger (總帳紀錄檔案)
    # cal_stt: the start date of the invoice matching
    # cal_end: the end date of the invoice matching
    def __init__(self, master, msgtxt):
        ttk.Frame.__init__(self, master)
        self.pack(side=TOP, fill=X)  # resized with parent
        title_label_font = font.Font(family="新細明體", size=12, weight="bold")
        self.msg = Label(self, wraplength='4i', justify=LEFT, font=title_label_font, pady=5)
        self.msg['text'] = ''.join(msgtxt)
        self.msg.pack(fill=X, padx=5, pady=5)
        self._fill_upper_frame()
        self._fill_middle_frame()
        self._fill_bottom_frame()

    # Upper frame constructed to accept the invoice issue record selected,
    # this selected invoice is of file type .xls
    def _fill_upper_frame(self):
        file_type = constant.ENTRY_TYPE_INVOICE_RECORD
        frame = ttk.Frame(self)
        lbl = ttk.Label(frame, width=20, text="統一發票開票紀錄", anchor='w')
        lbl.pack(side=LEFT)
        self.invoice_ent = ttk.Entry(frame, width=60)
        self.invoice_ent.pack(side=LEFT, expand=Y, padx=5)
        btn = ttk.Button(frame, text="選取",
                         command=lambda e=self.invoice_ent, t=file_type: self._file_dialog(e, t))
        btn.pack(side=LEFT, padx=5)
        frame.pack(side=TOP, padx='1c', pady=3)

    # Middle frame constructed to accept the general ledger selected,
    # this selected general ledger is of file type .xlsx
    def _fill_middle_frame(self):
        file_type = constant.ENTRY_TYPE_GENERAL_LEDGER
        frame = ttk.Frame(self)
        lbl = ttk.Label(frame, width=20, text="總帳紀錄", anchor='w')
        lbl.pack(side=LEFT)
        self.gl_ent = ttk.Entry(frame, width=60)
        self.gl_ent.pack(side=LEFT, expand=Y, padx=5)
        btn = ttk.Button(frame, text="選取",
                         command=lambda e=self.gl_ent, t=file_type: self._file_dialog(e, t))
        btn.pack(side=LEFT, padx=5)
        frame.pack(side=TOP, padx='1c', pady=3)

    # bottom frame includes two pairs of label and DateEntry objects to accept
    # start date and end date
    def _fill_bottom_frame(self):
        frame = ttk.Frame(self)
        # lbl_stt = ttk.Label(frame, width=15, text="發票起始時間")
        # lbl_stt.pack(side=LEFT)
        self.cal_stt_chk = IntVar()
        self.chkbtn_stt = Checkbutton(frame, text="發票起始日期", variable=self.cal_stt_chk, width=15,
                                      onvalue=1, offvalue=0)
        self.chkbtn_stt.pack(side=LEFT, padx=5)
        self.cal_stt = DateEntry(frame, width=25, background="darkblue",
                                 date_pattern="yyyy/MM/dd",
                                 foreground="white", borderwidth=2, year=2020)
        # self.cal_stt.delete(0, "end")
        self.cal_stt.pack(side=LEFT, padx=10)

        # lbl_end = ttk.Label(frame, width=15, text="發票結束時間")
        # lbl_end.pack(side=LEFT)
        self.cal_end_chk = IntVar()
        self.chkbtn_end = Checkbutton(frame, text="發票結束日期", variable=self.cal_end_chk, width=15,
                                      onvalue=1, offvalue=0)
        self.chkbtn_end.pack(side=LEFT, padx=5)
        self.cal_end = DateEntry(frame, width=20, background="darkblue",
                                 date_pattern="yyyy/MM/dd",
                                 foreground="white", borderwidth=2, year=2020)
        # self.cal_end.delete(0, "end")
        self.cal_end.pack(side=LEFT, padx=10)
        frame.pack(side=TOP, padx='1c', pady=3)

    # This is the file selector handler
    def _file_dialog(self, entry, file_type):
        fn = None
        op_pnl = self.master.op_pnl
        if file_type == constant.ENTRY_TYPE_GENERAL_LEDGER:
            opts = {"initialfile": entry.get(),
                    "filetypes": (('Excel Workbook', "*.xlsx"),
                                  ("All Files", "*.*"),),
                    "title": "選擇總帳紀錄檔"
                    }
            fn = fdlg.askopenfilename(**opts)
            op_pnl.print_log("選取總帳檔案: " + "\n\t" + fn)

        if file_type == constant.ENTRY_TYPE_INVOICE_RECORD:
            opts = {"initialfile": entry.get(),
                    "filetypes": (('Excel 97-2003 Workbook', "*.xls"),
                                  ("All Files", "*.*"),),
                    "title": "選擇統一發票開票紀錄檔"
                    }
            fn = fdlg.askopenfilename(**opts)
            op_pnl.print_log("選取發票檔案: "+ "\n\t" + fn)
        if fn:
            entry.delete(0, END)
            entry.insert(END, fn)


#
# class OperationPanel includes a log-text widget and three operation buttons
#
class OperationPanel(ttk.Frame):
    def __init__(self, master):
        ttk.Frame.__init__(self, master)
        self.pack(side=BOTTOM, fill=X)  # resize with parent

        # 'Match Invoice' button
        im = Image.open('..//images//compare.png')
        imh = ImageTk.PhotoImage(im)
        matchBtn = ttk.Button(text='比對銷貨紀錄', image=imh, default=ACTIVE, command=self.match_invoice)
        matchBtn.image = imh
        # configure button style
        matchBtn['compound'] = LEFT

        # 'Check Result' button
        im = Image.open('..//images//view.png')
        imh = ImageTk.PhotoImage(im)
        codeBtn = ttk.Button(text='檢視比對結果', image=imh, command=lambda: CodeDialog(self.master))
        codeBtn.image = imh
        codeBtn['compound'] = LEFT
        codeBtn.focus()

        # Dismiss button
        im = Image.open('..//images//exit.png')  # image file
        imh = ImageTk.PhotoImage(im)  # handle to file
        dismissBtn = ttk.Button(text='離開', image=imh, command=self.winfo_toplevel().destroy)
        dismissBtn.image = imh  # prevent image from being garbage collected
        dismissBtn['compound'] = LEFT  # display image to left of label text

        # separator widget
        # define customized font
        log_label_font = font.Font(family="新細明體", size=12, weight="bold")
        sep = ttk.Separator(orient=HORIZONTAL)
        log_label = ttk.Label(self, text="操作紀錄", justify=CENTER, font=log_label_font)

        # Log text frame, log_frame, which includes a vertical scrollbar and a log_text widget
        # Note that, in the pack method of log_text widget, 'expand' should be set to make the text box
        # extended in its parent frame
        log_frame = ttk.Frame(self)
        y_scrollbar = Scrollbar(log_frame, orient=VERTICAL)
        y_scrollbar.pack(side=RIGHT, fill=BOTH)
        self.log_text = Text(log_frame, yscrollcommand=y_scrollbar.set, spacing1=5, spacing2=3)
        self.log_text.pack(side=RIGHT, expand=1, fill=BOTH)
        y_scrollbar.config(command=self.log_text.yview)

        # position and register widgets as children of this frame
        sep.grid(in_=self, row=0, columnspan=5, sticky=EW, pady=10)
        log_label.grid(in_=self, row=1, columnspan=5, sticky=N, pady=5)
        # self.log_text.grid(in_=self, row=2, columnspan=5, sticky=EW, padx=10, pady=5)
        log_frame.grid(in_=self, row=2, columnspan=5, sticky=EW, padx=10, pady=5)
        matchBtn.grid(in_=self, row=3, column=0, sticky=E, padx=5, pady=10)
        codeBtn.grid(in_=self, row=3, column=1, sticky=E, padx=5, pady=10)
        dismissBtn.grid(in_=self, row=3, column=2, sticky=E, padx=5, pady=10)

        # set resize constraints
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        # bind <Return> to demo window, activates 'See Code' button;
        # <'Escape'> activates 'Dismiss' button
        self.winfo_toplevel().bind('<Return>', lambda x: codeBtn.invoke())
        self.winfo_toplevel().bind('<Escape>', lambda x: dismissBtn.invoke())

    # The button event handler when users click "比對銷貨紀錄" button
    def match_invoice(self):
        # sanity check of selected invoice records and general ledger
        inv_record_fn = self.master.sel_pnl.invoice_ent.get()
        gl_record_fn = self.master.sel_pnl.gl_ent.get()
        if inv_record_fn == "":
            self.print_log("尚未選擇發票檔案...")
            return
        if gl_record_fn == "":
            self.print_log("尚未選擇總帳檔案!")
            return
        self.print_log("執行發票、總帳匹配.....")
        self.print_log("發票檔案:" + inv_record_fn)
        self.print_log("總帳檔案:" + gl_record_fn)

        # self.master.sel_pnl.chkbtn_stt.select()
        chkbtn_stt = self.master.sel_pnl.cal_stt_chk.get()
        chkbtn_end = self.master.sel_pnl.cal_end_chk.get()
        if chkbtn_stt == 1:
            self.print_log("發票期始日 : " + self.master.sel_pnl.cal_stt.get())
        if chkbtn_end == 1:
            self.print_log("發票截止日 : " + self.master.sel_pnl.cal_end.get())

    def print_log(self, log_msg):
        now = datetime.now()
        time_stamp = now.strftime("[%Y/%m/%d %H:%M:%S] >> ")
        self.log_text.insert(END, time_stamp + log_msg + "\n")
        return
