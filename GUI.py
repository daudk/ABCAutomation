from tkinter import *
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

class IDSPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        ids_label = Label(self, text="IDS/S21/Oracle tasks",font=("Helvetica", 14))
        start_button = Button(self, text="Return to start page",
                                 command=lambda: master.switch_frame(StartPage))

        ids_label.grid(row=0,sticky=E+W+N+S,columnspan=2)



        btn_1 = Button(self, text = "Get Accounts Hitting 340",command = account_pulls_from_340,width=20)
        btn_1.grid(row=1,column=0,padx=10,pady=15)
        btn_1_label = Label(self, text = 'Given a sub and year, pulls all accounts hitting the "Payables" source for \
that sub and outputs them into new .csv file',wraplength=305,justify=LEFT)
        btn_1_label.grid(row=1,column=1,padx=10,sticky = W,pady=15)

        btn_2 = Button(self, text="Download All", command=download_all, width=20)
        btn_2.grid(row=2, column=0,padx=10)
        btn_2_label = Label(self, text='Download all reports from IDS that are available in .xlsx format.',
                            wraplength=305, justify=LEFT)
        btn_2_label.grid(row=2, column=1,pady=20,padx=10,sticky = W)




        start_button.grid(row=5,column=1,sticky=W+E+S,padx=(0,40),pady=15)

class SpdPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        spd_label = Label(self, text="Spreadsheet tasks", font=("Helvetica", 14))
        start_button = Button(self, text="Return to start page",
                                 command=lambda: master.switch_frame(StartPage))
        spd_label.grid(sticky='news',columnspan=2)

        btn_1 = Button(self, text="T_value Calculator", command=t_val_calc, width=20)
        btn_1.grid(row=1, column=0, padx=10)
        btn_1_label = Label(self, text='Point to folder containing TV6 files and the script will calculate T values from 8/2018 onwards and output summary file.', wraplength=305, justify=LEFT)
        btn_1_label.grid(row=1, column=1,pady=15, padx=10, sticky=W)

        btn_2 = Button(self, text="Extract Sample", command=sample_n, width=20)
        btn_2.grid(row=2, column=0, padx=10)
        btn_2_label = Label(self,
                            text='Extracts sample from total population and adds as a sheet to existing excel file.',
                            wraplength=305, justify=LEFT)
        btn_2_label.grid(row=2, column=1, pady=15, padx=10, sticky=W)

        btn_3 = Button(self, text="Rename All", command=rename_all, width=20)
        btn_3.grid(row=3, column=0, padx=10)
        btn_3_label = Label(self,
                            text='Renames download IDS reports to more meaningful names. (Currently only functional for 315)',
                            wraplength=305, justify=LEFT)
        btn_3_label.grid(row=3, column=1, pady=20, padx=10, sticky=W)

        btn_4 = Button(self, text="Combine Excel files", command=combine_files, width=20)
        btn_4.grid(row=4, column=0, padx=10)
        btn_4_label = Label(self, text='Combine multiple xlsx files into 1 file.', wraplength=305, justify=LEFT)
        btn_4_label.grid(row=4, column=1, pady=20, padx=10, sticky=W)

        start_button.grid(row=5, column=1, sticky=W + E + S, padx=(0, 40),pady=15)


class SOXPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        page_1_label = Label(self, text="This is the SOX Scoping page")
        start_button = Button(self, text="Return to start page",
                                 command=lambda: master.switch_frame(StartPage))
        page_1_label.grid()
        start_button.grid(pady=15)

class DevPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)

        page_1_label = Label(self, text="This is the Development page")
        start_button = Button(self, text="Return to start page",
                                 command=lambda: master.switch_frame(StartPage))
        page_1_label.grid()
        start_button.grid(pady=15)

#
# class PageTwo(Frame):
#     def __init__(self, master):
#         Frame.__init__(self, master)
#
#         page_2_label = Label(self, text="This is page two")
#         start_button = Button(self, text="Return to start page",
#                                  command=lambda: master.switch_frame(StartPage))
#         page_2_label.grid()
#         start_button.grid()


class StartPage(Frame):
    # Define settings upon initialization. Here you can specify
    def __init__(self, master=None):

        Frame.__init__(self, master)

        self.master = master


        # start_label = Label(self, text="This is the start page")
        # page_1_button = Button(self, text="Open page one",
        #                           command=lambda: master.switch_frame(IDSPage))
        # page_2_button = Button(self, text="Open page two",
        #                           command=lambda: master.switch_frame(PageTwo))
        # start_label.grid()
        # page_1_button.grid()
        # page_2_button.grid()

        btn_paths = "Resources/Buttons/"
        self.img_ids = PhotoImage(file=btn_paths + "ids_btn_1.png")
        self.img_sox = PhotoImage(file=btn_paths + "sox_btn_1.png")
        self.img_sps = PhotoImage(file=btn_paths + "sps_btn_1.png")
        self.img_dev = PhotoImage(file=btn_paths + "dev_btn_1.png")

        self.b_ids = Button(self, height=150, width=150, image=self.img_ids, command=lambda: master.switch_frame(IDSPage))
        self.b_ids.grid(row=1, column=1, padx=(70, 50), pady=10)

        self.b_sox = Button(self, height=150, width=150, image=self.img_sox, command=lambda: master.switch_frame(SOXPage))
        self.b_sox.grid(row=1, column=2, pady=10, padx=(0, 70))

        self.b_sps = Button(self, height=150, width=150, image=self.img_sps, command=lambda: master.switch_frame(SpdPage))
        self.b_sps.grid(row=2, column=1, padx=(70, 50), pady=5)

        self.b_dev = Button(self, height=150, width=150, image=self.img_dev, command=lambda: master.switch_frame(DevPage))
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
    #app.geometry("500x350")
    app.mainloop()
