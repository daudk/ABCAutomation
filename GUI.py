from tkinter import *
from tkinter import scrolledtext

from DEV_funcs import *
from IDS_funcs import *
from SPD_funcs import *


def client_exit():
    exit()


class SampleApp(Tk):
    def __init__(self):
        Tk.__init__(self)
        self._frame = None
        self.switch_frame(StartPage)

    def switch_frame(self, frame_class):
        """Destroys current frame and replaces it with a new one."""
        new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.grid()


class Pull315(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        edit_space_label = Label(self,text="Enter Account numbers separated by a new line: ",wraplength=200, justify=LEFT)
        edit_space = scrolledtext.ScrolledText(
            master=self,
            wrap='word',  # wrap text at full words only
            width=25,  # characters
            height=10,  # text lines
            bg='white'  # background color of edit area
        )
        edit_space.grid(row=1, column=0, padx=10,pady=(3,20), sticky=W,rowspan=3)
        edit_space_label.grid(row=0,column=0,sticky=W)


        subvar = StringVar(self)
        subvar.set("CUSA")
        sub = OptionMenu(self, subvar, *["CUSA", "CSA", "CFS", "CITS", "CIIS"])
        sub_label = Label(self, text='Subsidiary: ', wraplength=305, justify=LEFT)
        sub.config(width=7)
        sub.grid(row=1, column=2, padx=(2,30),pady=5)
        sub_label.grid(row=1, column=1, padx=2, sticky=W)

        yearvar = IntVar(self)
        yearvar.set(2017)
        year = OptionMenu(self, yearvar, *[2015, 2016, 2017, 2018])
        year_label = Label(self, text='Year: ', wraplength=305, justify=LEFT)
        year.config(width=7)
        year.grid(row=2, column=2, padx=(2,30),pady=5)
        year_label.grid(row=2, column=1, padx=2, sticky=W)


        freqvar = StringVar(self)
        freqvar.set("Yearly")
        freq = OptionMenu(self, freqvar, *["Yearly", "Biyearly", "Quarterly", "Monthly"])
        freq_label = Label(self, text='Frequency: ', wraplength=305, justify=RIGHT)
        freq.config(width=7)
        freq.grid(row=3, column=2, padx=(2,30),pady=(5,20),sticky='ew')
        freq_label.grid(row=3, column=1, padx=2, sticky=W)

        # from_mon_var = IntVar(self)
        # from_mon_var.set(1)
        # from_mon = OptionMenu(self, from_mon_var, *[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
        # from_mon.grid()
        #
        # to_mon_var = IntVar(self)
        # to_mon_var.set(1)
        # to_mon = OptionMenu(self, to_mon_var, *[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])
        # to_mon.grid()

        # go_button = Button(self, text="GO!", command=lambda: print_test(edit_space))


        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40))



class IDSPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        ids_label = Label(self, text="IDS tasks", font=("Helvetica", 14))
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        ids_label.grid(row=0, sticky=E + W + N + S, columnspan=2)

        btn_1 = Button(self, text="Get Accounts Hitting 340", command=account_pulls_from_340, width=20)
        btn_1.grid(row=1, column=0, padx=10)
        btn_1_label = Label(self, text='Given a sub and year, pulls all accounts hitting the "Payables" source for \
that sub and outputs them into new .csv file', wraplength=305, justify=LEFT)
        btn_1_label.grid(row=1, column=1, padx=10, sticky=W)

        btn_2 = Button(self, text="Download All", command=download_all, width=20)
        btn_2.grid(row=2, column=0, padx=10)
        btn_2_label = Label(self, text='Download all reports from IDS that are available in .xlsx format.',
                            wraplength=305, justify=LEFT)
        btn_2_label.grid(row=2, column=1, pady=20, padx=10, sticky=W)

        btn_3 = Button(self, text="Run 315 Reports", command=lambda: master.switch_frame(Pull315), width=20)
        btn_3.grid(row=3, column=0, padx=10)
        btn_3_label = Label(self, text='Runs 315 reports in bulk for multiple accounts.', wraplength=305, justify=LEFT)
        btn_3_label.grid(row=3, column=1, pady=20, padx=10, sticky=W)

        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40))


class SpdPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        spd_label = Label(self, text="Spreadsheet Tasks", font=("Helvetica", 14))
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        spd_label.grid(row=0, sticky=E + W + N + S, columnspan=2)

        btn_1 = Button(self, text="T_value Calculator", command=t_val_calc, width=20)
        btn_1.grid(row=1, column=0, padx=10, pady=15)
        btn_1_label = Label(self,
                            text='Point to folder containing TV6 files and the script will calculate T values from 8/2018 onwards and output summary file.',
                            wraplength=305, justify=LEFT)
        btn_1_label.grid(row=1, column=1, padx=10, sticky=W, pady=15)

        btn_2 = Button(self, text="Extract Sample", command=sample_n, width=20)
        btn_2.grid(row=2, column=0, padx=10, pady=20)
        btn_2_label = Label(self,
                            text='Extracts sample from total population and adds as a sheet to existing excel file.',
                            wraplength=305, justify=LEFT)
        btn_2_label.grid(row=2, column=1, padx=10, sticky=W, pady=20)

        btn_3 = Button(self, text="Rename All", command=rename_all, width=20)
        btn_3.grid(row=3, column=0, padx=10, pady=20)
        btn_3_label = Label(self,
                            text='Renames download IDS reports to more meaningful names. (Currently only functional for 315)',
                            wraplength=305, justify=LEFT)
        btn_3_label.grid(row=3, column=1, pady=20, padx=10, sticky=W)

        btn_4 = Button(self, text="Combine Excel files", command=combine_files, width=20)
        btn_4.grid(row=4, column=0, padx=10)
        btn_4_label = Label(self, text='Combine multiple xlsx files into 1 file.', wraplength=305, justify=LEFT)
        btn_4_label.grid(row=4, column=1, pady=20, padx=10, sticky=W)

        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40), pady=20)


class SOXPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        page_1_label = Label(self, text="This is the SOX Scoping page")
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))
        page_1_label.grid(row=0, sticky=E + W + N + S, columnspan=2)

        start_button.grid()


class DevPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        dev_label = Label(self, text="Tasks still under development", font=("Helvetica", 14))
        dev_label_2 = Label(self, text="Don't use this stuff, will probably cause problems.", font=("Helvetica", 10))
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        dev_label.grid(row=0, sticky=E + W + N + S, columnspan=2)
        dev_label_2.grid(row=1, sticky=E + W + N + S, columnspan=2)
        start_button = Button(self, text="Return to start page",
                              command=lambda: master.switch_frame(StartPage))

        btn_1 = Button(self, text="Pull S21 Invoices", command=s21_invoices, width=20)
        btn_1.grid(row=2, column=0, padx=10, pady=15)
        btn_1_label = Label(self,
                            text='Pulls a list of invoices from S21',
                            wraplength=305, justify=LEFT)
        btn_1_label.grid(row=2, column=1, padx=10, sticky=W, pady=15)

        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40), pady=20)


class StartPage(Frame):
    # Define settings upon initialization. Here you can specify
    def __init__(self, master=None):
        Frame.__init__(self, master)

        self.master = master

        btn_paths = "Resources/Buttons/"
        self.img_ids = PhotoImage(file=btn_paths + "ids_btn_1.png")
        self.img_sox = PhotoImage(file=btn_paths + "sox_btn_1.png")
        self.img_sps = PhotoImage(file=btn_paths + "sps_btn_1.png")
        self.img_dev = PhotoImage(file=btn_paths + "dev_btn_1.png")

        self.b_ids = Button(self, height=150, width=150, image=self.img_ids,
                            command=lambda: master.switch_frame(IDSPage))
        self.b_ids.grid(row=1, column=1, padx=(70, 50), pady=10)

        self.b_sox = Button(self, height=150, width=150, image=self.img_sox,
                            command=lambda: master.switch_frame(SOXPage))
        self.b_sox.grid(row=1, column=2, pady=10, padx=(0, 70))

        self.b_sps = Button(self, height=150, width=150, image=self.img_sps,
                            command=lambda: master.switch_frame(SpdPage))
        self.b_sps.grid(row=2, column=1, padx=(70, 50), pady=5)

        self.b_dev = Button(self, height=150, width=150, image=self.img_dev,
                            command=lambda: master.switch_frame(DevPage))
        self.b_dev.grid(row=2, column=2, pady=5, padx=(0, 70))

        self.init_window()

    # Creation of init_window
    def init_window(self):
        # changing the title of our master widget
        self.master.title("ABC Automation Platform")

        # allowing the widget to take the full space of the root window
        # self.pack(fill=BOTH, expand=1)
        self.grid()

        # creating a menu instance
        menu = Menu(self)
        # self.master.config(menu=menu)

        # create the file object)
        file = Menu(menu, tearoff=False)
        file.add_command(label="Exit", command=client_exit)

        file.add_command(label="Get Accounts To Pull", command=account_pulls_from_340)
        file.add_command(label="Download All", command=download_all)
        file.add_command(label="Rename All", command=rename_all)
        menu.add_cascade(label="File", menu=file)

        edit = Menu(menu, tearoff=False)
        help = Menu(menu, tearoff=False)
        help.add_command(label="Help")
        edit.add_command(label="Undo")
        menu.add_cascade(label="Edit", menu=edit)
        menu.add_cascade(label="Help", menu=help)

        self.master.config(menu=menu)


if __name__ == "__main__":
    app = SampleApp()
    # app.geometry("500x350")
    app.mainloop()
