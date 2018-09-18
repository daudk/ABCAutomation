from os.path import expanduser
from GUI import *
from IDS_funcs import *
from SOX_funcs import *
from tkinter import scrolledtext



new_directory=""

def print_out(msg):
    print (msg)



class SampleSelection(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        ctrlVar = StringVar(self)
        ctrlVar.set("")
        ctrl = OptionMenu(self,ctrlVar,*["15.2.3.02 F&A","15.2.1.08 BISG","15.2.1.10 ITCG","15.2.1.07 LAG","15.2.1.16 BISG",
                                         "15.2.1.13 BISG","15.2.1.13 LAG","15.2.1.12 ITCG","15.2.6.09 ITCG","4.6.09 ITCG"])
        ctrl_label=Label(self, text="Control Number: ", wraplength=305,justify=LEFT)
        ctrl.config(width=10)
        ctrl.grid(row=1, column=1, pady=5,sticky=W)
        ctrl_label.grid(row=1, column=0, padx=5, sticky=W)

        samp_size_var =  StringVar()
        samp_size = Entry(self,textvariable=samp_size_var)
        samp_size_var.set("25")
        samp_size.config(width=3)
        samp_size.grid(row=2,column=1,sticky=W)
        samp_size_label = Label(self, text ="# of Samples: ")
        samp_size_label.grid(row=2, column=0, padx=5, sticky=W)



        text = Text(self, state='disabled', width=35, height=1)
        text.configure(state='disabled')

        browse_button = Button(self, text="Browse", command=lambda: file_browser("Point to Manage Expense Pickle File!"))
        download_label = Label(self, text="Master File: ", justify=LEFT)
        download_label.grid(row=3, column=0,padx=(5, 5), sticky=W)

        def file_browser(msg):
            new_directory = filedialog.askopenfilename(title=msg,initialdir = "/",filetypes = [("Pickle Files","*.pkl")])
            text.configure(state='normal')
            text.delete(1.0, 'end')
            text.insert(1.0, new_directory)
            text.configure(state='disabled')

        go_button = Button(self, text="GO!",
                           command=lambda: sox_test_samples(ctrlVar.get(),samp_size_var.get(),text.get(1.0,'end').strip()))
        text.grid(row=3, column=1,columnspan=2)
        browse_button.grid(row=3, column=3, padx=15, pady=10)

        start_button.grid(row=5, column=0, padx=(15, 0))
        go_button.grid(row=5,column=3,sticky=W+E+N+S,padx=(0,15))

class Download(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        dl_path = StringVar()
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))
        download_label = Label(self, text="Download Directory: ", justify=LEFT)
        download_label.grid(row=1, column=0, sticky=W)
        new_directory = expanduser("~") + "\\Downloads\\"

        text = Text(self, state='disabled', width=50, height=1)
        text.configure(state='normal')
        text.insert('end', new_directory )
        text.configure(state='disabled')

        browse_button = Button(self, text="Browse", command=lambda: file_browser("Point to Download directory!"))



        def file_browser(msg):
            new_directory  = filedialog.askdirectory(title=msg)
            text.configure(state='normal')
            text.delete(1.0,'end')
            text.insert(1.0,new_directory)
            text.configure(state='disabled')

        go_button = Button(self, text="GO!",
                           command=lambda: download_all(text.get(1.0,END)))

        text.grid(row=1, column=1)
        browse_button.grid(row=1, column=2,padx=(10,15), pady=(15,10))

        go_button.grid(row=3,column=2,sticky=E+W+S+N, padx=(10,15))
        start_button.grid(row=3,column=0,padx=(15,0))


class InvoicePulls(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        directory = expanduser("~") + "\\Downloads\\"

        edit_space_label = Label(self, text="Enter Invoice numbers separated by a new line: ", wraplength=200,
                                 justify=LEFT)
        edit_space = scrolledtext.ScrolledText(
            master=self,
            wrap='word',  # wrap text at full words only
            width=25,  # characters
            height=10,  # text lines
            bg='white'  # background color of edit area
        )
        edit_space.grid(row=1, column=0, padx=10, pady=(3, 20), sticky=W, rowspan=5)
        edit_space_label.grid(row=0, column=0, sticky=W)



        text = Text(self, state='disabled', width=28, height=3)
        text.configure(state='normal')
        text.insert('end', directory)
        text.configure(state='disabled')

        browse_button = Button(self, text="Browse", command=lambda: file_browser("Point to a new empty directory!"))


        def file_browser(msg):
            directory = filedialog.askdirectory(title=msg)+"/"
            text.configure(state='normal')
            text.delete(1.0, 'end')
            text.insert('end', directory)
            text.configure(state='disabled')


        Label(self,text="Browse to a download drectory: ").grid(row=2,column=1,sticky=W+S,pady=(0,0))
        text.grid(row=3,column=1,columnspan=2,sticky=W+N)
        browse_button.grid(row=3,column=3,sticky=E,padx=(5,12))
        go_button = Button(self, text="GO!",
                           command=lambda: dl_s21_invoices(text.get(1.0,'end'),edit_space))

        go_button.grid(row=5, column=2, sticky=W + E + S, padx=(0, 40))
        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40))

class Pull315(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        edit_space_label = Label(self, text="Enter Account numbers separated by a new line: ", wraplength=200,
                                 justify=LEFT)
        edit_space = scrolledtext.ScrolledText(
            master=self,
            wrap='word',  # wrap text at full words only
            width=25,  # characters
            height=10,  # text lines
            bg='white'  # background color of edit area
        )
        edit_space.grid(row=1, column=0, padx=10, pady=(3, 20), sticky=W, rowspan=5)
        edit_space_label.grid(row=0, column=0, sticky=W)

        # sources=

        sourcevar = StringVar(self)
        sourcevar.set("ALL VALUES")
        source = OptionMenu(self, sourcevar,
                            *['ALL VALUES', 'ACCTS RECEIVABLE', 'Assets', 'CLA Reclass', 'CNA', 'CONVERSION',
                              'COST CALCULATION',
                              'CPNet', 'DHPS', 'EMPLOYEE REIMB', 'ESS3', 'EXPORT A/R', 'EXPORT SALES', 'IPO', 'LOAN',
                              'MA TRANSFER',
                              'MDSE INVENTORY', 'Manual', 'MassAllocation', 'NETWORK MGMT', 'NTC', 'Other', 'PARTS',
                              'PAYROLL',
                              'Payables', 'RENTAL ASSETS', 'ROSS', 'Recurring', 'Revaluation', 'Ross', 'S21 A/R',
                              'S21 Accounting AR',
                              'S21 Accounting Cost', 'S21 Cost', 'S21 Export', 'S21 Logistics', 'S21 Order',
                              'S21 Parts',
                              'S21 Parts (Export)', 'S21 Parts Export', 'S21 Procurement', 'Spreadsheet', 'TRANZACT',
                              'WHOLESALE',
                              'Web-ADI', 'eStore'])
        source_label = Label(self, text="Source: ", wraplength=305, justify=LEFT)
        source.config(width=18)
        source.grid(row=4, column=2, padx=(0, 30), pady=(5, 20), sticky='ew')
        source_label.grid(row=4, column=1, padx=2, sticky=W)

        subvar = StringVar(self)
        subvar.set("CUSA")
        sub = OptionMenu(self, subvar, *["CUSA", "CFS", "CITS", "CCI","CIS", "CPA"])
        sub_label = Label(self, text='Subsidiary: ', wraplength=305, justify=LEFT)
        sub.config(width=7)
        sub.grid(row=1, column=2, padx=(2, 5), pady=5)
        sub_label.grid(row=1, column=1, padx=2, sticky=W)

        yearvar = IntVar(self)
        yearvar.set(2017)
        year = OptionMenu(self, yearvar, *[2015, 2016, 2017, 2018])
        year_label = Label(self, text='Year: ', wraplength=305, justify=LEFT)
        year.config(width=7)
        year.grid(row=2, column=2, padx=(2, 5), pady=5)
        year_label.grid(row=2, column=1, padx=2, sticky=W)

        freqvar = StringVar(self)
        freqvar.set("Yearly")
        freq = OptionMenu(self, freqvar, *["Yearly", "Semi-annually", "Quarterly", "Monthly"])
        freq_label = Label(self, text='Frequency: ', wraplength=305, justify=RIGHT)
        freq.config(width=7)
        freq.grid(row=3, column=2, padx=(2, 30), pady=(5, 5), sticky='ew')
        freq_label.grid(row=3, column=1, padx=2, sticky=W)

        go_button = Button(self, text="GO!",
                           command=lambda: parse_315_input(edit_space, subvar, yearvar, freqvar, sourcevar))

        go_button.grid(row=5, column=2, sticky=W + E + S, padx=(0, 40))
        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40))
