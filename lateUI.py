import tkinter as tk
import tkinter.messagebox
from tkinter.filedialog import askopenfile
import xlrd
from tkinter import filedialog
import os
import pyodbc
import datetime


class ui_frame:
    def __init__(self):
        self.window = tk.Tk()
        # background grey color
        self.bg = "#d2d4dc"
        self.bg_title = "#4c516d"
        self.window.config(bg=self.bg)
        self.frame_a = tk.Frame(bg=self.bg)
        self.frame_title = tk.Frame(bg=self.bg, padx=30)
        self.frame_b = tk.Frame()
        self.frame_b.config(bg=self.bg, padx=20)
        self.photo = tk.PhotoImage(file=r"kk.png")
        self.cross_box = tk.PhotoImage(file=r"crossButton.png")

        self.frameCheck = tk.Frame()
        self.frameCheck.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck2 = tk.Frame()
        self.frameCheck2.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck3 = tk.Frame()
        self.frameCheck3.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck4 = tk.Frame()
        self.frameCheck4.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck5 = tk.Frame()
        self.frameCheck5.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck6 = tk.Frame()
        self.frameCheck6.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck7 = tk.Frame()
        self.frameCheck7.config(bg=self.bg, padx=15, pady=3)
        self.frameCheck8 = tk.Frame()
        self.frameCheck8.config(bg=self.bg, padx=15, pady=3)

        self.framebutton = tk.Frame()
        self.framebutton.config(bg=self.bg, pady=7)

        self.meta_check_1 = tk.IntVar()
        self.meta_check_2 = tk.IntVar()
        self.integrity_check_1 = tk.IntVar()
        self.integrity_check_2 = tk.IntVar()
        self.quality_check = tk.IntVar()
        self.completeness_check = tk.IntVar()
        self.cleansing_check = tk.IntVar()
        self.transformation_check = tk.IntVar()
        self.xml_data_load_check = tk.IntVar()
        self.stored_proc_check = tk.IntVar()

        self.conn = pyodbc.connect("Driver={ODBC Driver 17 for SQL Server};"
                                   "Server=CDC2-L-F1TGQEH;"
                                   "Database=test;"
                                   "Trusted_Connection=yes;"
                                   )
        self.cursor = self.conn.cursor()

    def a(self):
        # Heading-blue
        logo = tk.PhotoImage(file=r"new.png")
        label_a = tk.Label(master=self.frame_a, image=logo)
        label_a.image = logo
        label_a.pack(pady=(0, 20))
        label_a.config(font=("Calibri", "24", "bold"))

        # SubHeading-grey
        label_b = tk.Label(self.frame_title, text="Test Types", fg="white", bg=self.bg_title, anchor="w")
        label_b.config(font=("Calibri", "12", "bold"))
        label_b.grid(row=3, column=1, pady=15, sticky="w", ipadx=46, ipady=5)
        label_b = tk.Label(self.frame_title, text="Test Objective", fg="white", bg=self.bg_title, width=72, anchor="w")
        label_b.config(font=("Calibri", "12", "bold"))
        label_b.grid(row=3, column=2, padx=20, ipady=5, ipadx=89, sticky="w")
        label_b = tk.Label(self.frame_title, text="Applicable Test Phase ", fg="white", bg=self.bg_title, anchor="w")
        label_b.config(font=("Calibri", "12", "bold"))
        label_b.grid(row=3, column=3, sticky="w", ipady=5, padx=10, ipadx=13)
        label_a.pack()

    def b(self):

        #  ROW ONE
        label_b = tk.Label(self.frame_b, text="MetaData Test", fg="white", bg="#0057e7", activebackground='#0057e7')
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=1, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate target tables are built as per the source table schema definition/data definition model",
                           fg="black", bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=1, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW TWO
        label_b = tk.Label(self.frame_b, text="Data Integrity Test", fg="white", bg="#0057e7",
                           activebackground='#0057e7')
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=5, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate that the referential integrity between the tables are established accurately",
                           fg="black", bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=5, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW THREE
        label_b = tk.Label(self.frame_b, text="Data Quality Test", fg="white", bg="#0057e7")
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=6, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate that the content/records of the table are as per the data type and constraint definition of the element",
                           fg="black", bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=6, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW F0UR
        label_b = tk.Label(self.frame_b, text="Data Completeness Test", fg="white", bg="#0057e7")
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=7, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b, text="To validate the accuracy data load from source to target tables",
                           fg="black",
                           bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=7, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW FIVE
        label_b = tk.Label(self.frame_b, text="Data Cleansing Test", fg="white", bg="#0057e7")
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=8, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate that the elements of target table against the data cleansing rules",
                           fg="black",
                           bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=8, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW SIX
        label_b = tk.Label(self.frame_b, text="Data Transformation Test", fg="white", bg="#0057e7")
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=9, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate that data transformed into IDS is as per the business transoformation logic",
                           fg="black",
                           bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=9, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW SEVEN
        label_b = tk.Label(self.frame_b, text="Hop2 - XML Data Load Test", fg="white", bg="#0057e7")
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=10, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate that the xml data in IDS Staging is loaded appropriately into IDS table as per the XSD definition.",
                           fg="black",
                           bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=10, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

        # ROW EIGHT
        label_b = tk.Label(self.frame_b, text="Hop3 - Stored Proc Test", fg="white", bg="#0057e7")
        label_b.config(font=("Calibri", "11", "bold"))
        label_b.grid(row=11, column=3, padx=10, pady=10, ipady=7, sticky="e,w")
        label_b = tk.Label(self.frame_b,
                           text="To validate that the xml generated by the Hop3 Stored Proc is accurate and has all requried IDS data based on extract query",
                           fg="black",
                           bg="white", anchor="w")
        label_b.config(font=("Calibri", "11"))
        label_b.grid(row=11, column=4, padx=10, pady=10, ipady=7, sticky="e,w")

    def checkbox_frame1(self):

        label_b = tk.Label(self.frameCheck, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")

        c = tk.Checkbutton(self.frameCheck, bg=self.bg, variable=self.meta_check_1)
        c.grid(row=0, column=2, sticky="e,w")

        label_b = tk.Label(self.frameCheck, text="   Hop2", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck, bg=self.bg, variable=self.meta_check_2)
        c.grid(row=0, column=4, padx=2, ipady=2)

    def checkbox_frame2(self):
        label_b = tk.Label(self.frameCheck2, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")

        c = tk.Checkbutton(self.frameCheck2, bg=self.bg, variable=self.integrity_check_1)
        c.grid(row=0, column=2, sticky="e,w")

        label_b = tk.Label(self.frameCheck2, text="   Hop2", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck2, bg=self.bg, variable=self.integrity_check_2)
        c.grid(row=0, column=4, padx=2, ipady=2)

    def checkbox_frame3(self):
        label_b = tk.Label(self.frameCheck3, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")
        label_image_cross = tk.Label(self.frameCheck3, text="\u274e", bg=self.bg)
        label_image_cross.grid(row=0, column=2)
        label_b = tk.Label(self.frameCheck3, text="     Hop2 ", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck3, bg=self.bg, variable=self.quality_check)
        c.grid(row=0, column=4, sticky="e,w", ipady=2)

    def checkbox_frame4(self):
        label_b = tk.Label(self.frameCheck4, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck4, bg=self.bg, variable=self.completeness_check)
        c.grid(row=0, column=2, sticky="e,w")
        label_b = tk.Label(self.frameCheck4, text="   Hop2", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        label_image_cross = tk.Label(self.frameCheck4, text="\u274e", bg=self.bg)
        label_image_cross.grid(row=0, column=4, padx=4.5, ipady=5)

    def checkbox_frame5(self):
        label_b = tk.Label(self.frameCheck5, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")
        label_image_cross = tk.Label(self.frameCheck5, text="\u274e", bg=self.bg)
        label_image_cross.grid(row=0, column=2)
        label_b = tk.Label(self.frameCheck5, text="     Hop2 ", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck5, bg=self.bg, variable=self.cleansing_check)
        c.grid(row=0, column=4, sticky="e,w", pady=2, ipady=2)

    def checkbox_frame6(self):
        label_b = tk.Label(self.frameCheck6, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")
        label_image_cross = tk.Label(self.frameCheck6, text="\u274e", bg=self.bg)
        label_image_cross.grid(row=0, column=2)
        label_b = tk.Label(self.frameCheck6, text="     Hop2 ", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck6, bg=self.bg, variable=self.transformation_check)
        c.grid(row=0, column=4, sticky="e,w", pady=2)

    def checkbox_frame7(self):
        label_b = tk.Label(self.frameCheck7, text="Hop1", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1, sticky="e,w")
        label_image_cross = tk.Label(self.frameCheck7, text="\u274e", bg=self.bg)
        label_image_cross.grid(row=0, column=2)
        label_b = tk.Label(self.frameCheck7, text="     Hop2 ", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=3, sticky="e,w")
        c = tk.Checkbutton(self.frameCheck7, bg=self.bg, variable=self.xml_data_load_check)
        c.grid(row=0, column=4, sticky="e,w", pady=2)

    def checkbox_frame8(self):
        label_b = tk.Label(self.frameCheck8, text="          Hop3", bg=self.bg)
        label_b.config(font=("", "10", "bold"))
        label_b.grid(row=0, column=1)
        c = tk.Checkbutton(self.frameCheck8, bg=self.bg, variable=self.stored_proc_check)
        c.grid(row=0, column=4, pady=2)

    def buttons_frame(self):
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.meta_data_test)
        label_b_button.grid(row=1, column=1, padx=10, pady=2)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.data_integrity_test)
        label_b_button.grid(row=2, column=1, padx=10, pady=11)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.data_quality_test)
        label_b_button.grid(row=3, column=1, padx=10, pady=1)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.data_completeness_test)
        label_b_button.grid(row=4, column=1, padx=10, pady=11)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.data_cleansing_test)
        label_b_button.grid(row=5, column=1, padx=10, pady=3)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg)
        label_b_button.grid(row=6, column=1, padx=10, pady=8)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.hop2_xml_data_load_test)
        label_b_button.grid(row=7, column=1, padx=10, pady=4.5)
        label_b_button = tk.Button(self.framebutton, text='Click Me !', image=self.photo, relief="flat", bg=self.bg,
                                   command=self.hop3_stored_proc_test)
        label_b_button.grid(row=8, column=1, padx=10, pady=5)

    def meta_data_test(self):

        if self.meta_check_1.get() == self.meta_check_2.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)
            if self.meta_check_1.get() == self.meta_check_2.get() == 1:
                print(" Both hop is selected ")
                if "Hop1_MetaData_Input" in wb.sheet_names() and 'Hop2_MetaData_Input' in wb.sheet_names():
                    print("present")
                    sheet = wb.sheet_by_name('Hop1_MetaData_Input')
                    print(sheet.nrows)
                    sheet_col = sheet.ncols
                    sheet2 = wb.sheet_by_name('Hop2_MetaData_Input')
                    print(sheet2.nrows)
                    sheet2_col = sheet2.ncols
                    if sheet.nrows >= 2:
                        if sheet2.nrows >= 2:
                            print(" can starttttttt hop2 and hop2")
                            # Starting to check the cells

                            for i in range(sheet_col):
                                try:
                                    c = sheet.cell_type(1, i)
                                    print(c)
                                    if c == 0:
                                        tk.messagebox.showinfo("eroor", "Metadat-Hop1 Input sheet have empty values")
                                        print(" empty cell")
                                    else:
                                        continue
                                except:
                                    c = 0
                            print("sheet2")
                            for j in range(sheet2_col):
                                try:
                                    c2 = sheet2.cell_type(1, j)
                                    print(c2)
                                    if c2 == 0:
                                        tk.messagebox.showinfo("eroor", "Metadat-Hop2 Input sheet has a empty values ")
                                        print("empty cell")
                                    else:
                                        continue
                                except:
                                    pass

                        else:
                            tk.messagebox.showinfo("Important Message",
                                                   "Metadat-Hop2 Input sheet does not have any input to run the test. ")
                    else:
                        tk.messagebox.showinfo("Important Message",
                                               "Metadat-Hop1 Input sheet does not have any input to run the test. ")
                else:
                    print("not present")
                    tk.messagebox.showinfo("Important Message", "Input File doesn't have the required Sheets"
                                                                "\n File Location : " + filename)
                    breakpoint()

            elif self.meta_check_1.get() == 1:
                print("hop 1 is selected with variable " + str(self.meta_check_1.get()))
                if "Hop1_MetaData_Input" in wb.sheet_names():
                    sheet_Hop1 = wb.sheet_by_name('Hop1_MetaData_Input')

                    if sheet_Hop1.nrows >= 2:
                        print("proceed")

                        for i in range(sheet_Hop1.ncols):
                            try:
                                c = sheet_Hop1.cell_type(1, i)
                                print(c)
                                if c == 0:
                                    tk.messagebox.showinfo("Error", "Metadata-Hop1 Input sheet have empty values")
                                    print(" empty cell")
                                else:
                                    continue
                            except:
                                c = 0
                    else:
                        tk.messagebox.showinfo("Important Message",
                                               "Metadata-Hop1 Input sheet does not have any input to run the test. ")
                else:
                    print("not present")
                    tk.messagebox.showinfo("Important Message", "Input File doesn't have the required Sheets"
                                                                "\n File Location : " + filename)
                    breakpoint()


            else:
                print("hop 2 is selected withe variable " + str(self.meta_check_2.get()))
                if "Hop2_MetaData_Input" in wb.sheet_names():

                    sheet_Hop2 = wb.sheet_by_name('Hop2_MetaData_Input')
                    if sheet_Hop2.nrows >= 2:
                        print("proceed 2 ")
                        for j in range(sheet_Hop2.ncols):
                            try:
                                c2 = sheet_Hop2.cell_type(1, j)
                                print(c2)
                                if c2 == 0:
                                    tk.messagebox.showinfo("Error", "Metadata-Hop2 Input sheet has a empty values ")
                                    print("empty cell")
                                else:
                                    continue
                            except:
                                pass
                    else:
                        tk.messagebox.showinfo("Important Message",
                                               " Metadata-Hop2 Input sheet does not have any input to run the test. ")
                else:
                    print("not present")
                    tk.messagebox.showinfo("Important Message", "Input File doesn't have the required Sheets"
                                                                "\n File Location : " + filename)
                    breakpoint()

    def data_integrity_test(self):
        if self.integrity_check_1.get() == self.integrity_check_2.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)
            if self.integrity_check_1.get() == self.integrity_check_2.get() == 1:
                print(" Both hop is selected ")
                if "Hop1_DataIntegrity_Input" in wb.sheet_names() and 'Hop2_DataIntegrity_Input' in wb.sheet_names():
                    print("present")
                    sheet = wb.sheet_by_name('Hop1_DataIntegrity_Input')
                    print(sheet.nrows)
                    sheet_col = sheet.ncols
                    sheet2 = wb.sheet_by_name('Hop2_DataIntegrity_Input')
                    print(sheet2.nrows)
                    sheet2_col = sheet2.ncols
                    if sheet.nrows >= 2:
                        if sheet2.nrows >= 2:
                            print(" can starttttttt hop2 and hop2")
                            # Starting to check the cells

                            for i in range(sheet_col - 2):
                                try:
                                    c = sheet.cell_type(1, i)
                                    print(c)
                                    if c == 0:
                                        tk.messagebox.showinfo("error",
                                                               "DataIntegrity-Hop1 Input sheet have empty values")
                                        print(" empty cell")
                                    else:
                                        continue
                                except:
                                    c = 0
                            print("sheet2")
                            for j in range(sheet2_col - 2):
                                try:
                                    c2 = sheet2.cell_type(1, j)
                                    print(c2)
                                    if c2 == 0:
                                        tk.messagebox.showinfo("error",
                                                               "DataIntegrity-Hop2 Input sheet has a empty values ")
                                        print("empty cell")
                                    else:
                                        continue
                                except:
                                    pass

                        else:
                            tk.messagebox.showinfo("Important Message",
                                                   "DataIntegrity-Hop2 Input sheet does not have any input to run the test. ")
                    else:
                        tk.messagebox.showinfo("Important Message",
                                               "DataIntegrity-Hop1 Input sheet does not have any input to run the test. ")
                else:
                    print("not present")
                    tk.messagebox.showinfo("Important Message", "Input File doesn't have the required Sheets"
                                                                "\n File Location : " + filename)
                    breakpoint()

            elif self.integrity_check_1.get() == 1:
                print("hop 1 is selected with variable " + str(self.integrity_check_1.get()))
                sheet_Hop1 = wb.sheet_by_name('Hop1_DataIntegrity_Input')
                if sheet_Hop1.nrows >= 2:
                    print("proceed")
                    for i in range(sheet_Hop1.ncols - 2):
                        try:
                            c = sheet_Hop1.cell_type(1, i)
                            print(c)
                            if c == 0:
                                tk.messagebox.showinfo("error", "DataIntegrity-Hop1 Input sheet have empty values")
                                print(" empty cell")
                            else:
                                continue
                        except:
                            c = 0
                else:
                    tk.messagebox.showinfo("Important Message",
                                           "DataIntegrity-Hop1 Input sheet does not have any input to run the test. ")
            else:
                print("hop 2 is selected withe variable " + str(self.integrity_check_2.get()))
                sheet_Hop2 = wb.sheet_by_name('Hop2_DataIntegrity_Input')
                if sheet_Hop2.nrows >= 2:
                    print("proceed 2 ")
                    for j in range(sheet_Hop2.ncols - 2):
                        try:
                            c2 = sheet_Hop2.cell_type(1, j)
                            print(c2)
                            if c2 == 0:
                                tk.messagebox.showinfo("error", "DataIntegrity-Hop2 Input sheet has a empty values ")
                                print("empty cell")
                            else:
                                continue
                        except:
                            pass
                else:
                    tk.messagebox.showinfo("Important Message",
                                           " DataIntegrity-Hop2 Input sheet does not have any input to run the test. ")

    def data_quality_test(self):
        if self.quality_check.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)

            if self.quality_check.get() == 1:
                print("hop 2 is selected with variable " + str(self.integrity_check_1.get()))

                if 'Hop2_DataQuality_Input' in wb.sheet_names():
                    sheet_hop2 = wb.sheet_by_name('Hop2_DataQuality_Input')
                    if sheet_hop2.nrows > 2:
                        for r in range(1, sheet_hop2.nrows):
                            for c in range(sheet_hop2.ncols - 2):
                                cells_present_flag = 1
                                try:
                                    c = sheet_hop2.cell_value(r, c)
                                    if c:
                                        # print("present" + r)
                                        print("")
                                    else:
                                        tk.messagebox.showinfo("Error", "DataQuality-Hop2 Input sheet have empty cells")
                                        cells_present_flag = 0
                                        break
                                except Exception as e:
                                    print(e)
                            if cells_present_flag == 1:
                                t = sheet_hop2.cell_value(r, 2)
                                dates = []
                                rating = []
                                null_rows = []
                                employee_name = []
                                self.cursor.execute('select * from {}'.format(t))
                                for row in self.cursor:
                                    # print(row)
                                    dates.append(row[2])
                                    employee_name.append(row[3])
                                    rating.append(row[4])
                                    # null row check
                                    for c in row:
                                        # print(c)
                                        if c is None:
                                            null_rows.append(1)
                                            break
                                # print(null_rows)
                                # print(dates)
                                # print(rating)
                                # print(employee_name)
                                # reading the validation check from the column 6
                                choose_Validation_check = sheet_hop2.cell_value(r, 5)
                                if choose_Validation_check == 'Null Check':
                                    print("null")
                                    print(null_rows)
                                elif choose_Validation_check == 'Data Truncation Check':
                                    print("trunc")
                                    mk = 0
                                    for e in employee_name:
                                        e_len = len(e)
                                        print(e, len(e))
                                        if e_len > mk:
                                            mk = e_len
                                    print(mk)
                                elif choose_Validation_check == 'Duplicate Check':
                                    print("dup")
                                elif choose_Validation_check == "Num Field Format":
                                    print("numm")
                                    num_v_check = ""
                                    for rate in rating:
                                        if type(rate) is int:
                                            num_v_check = "pass"
                                        else:
                                            num_v_check = "fail"
                                            break
                                elif choose_Validation_check == "Date Field Format":
                                    print("date")
                                    date_v_check = ""
                                    for date in dates:
                                        if type(date) is datetime.date:
                                            date_v_check = "pass"
                                        else:
                                            date_v_check = "fail"
                                            break
                                else:
                                    pass



                            else:
                                print("ganths")
                            print("\n")
                            # print(null_rows)
                    else:
                        tk.messagebox.showinfo("Error", "Hop2 : input sheet does not have the inputs to run the test")
                else:
                    tk.messagebox.showinfo("Error", "Input File does not have the required sheet"
                                                    "\n File Location : " + filename)

    def data_completeness_test(self):
        if self.completeness_check.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)

            if self.completeness_check.get() == 1:
                print("hop 1 is selected with variable " + str(self.integrity_check_1.get()))
                sheet_Hop1 = wb.sheet_by_name('Hop1_DataCompleteness_Input')
                if sheet_Hop1.nrows >= 2:
                    print("proceed")
                    for i in range(sheet_Hop1.ncols - 3):
                        try:
                            c = sheet_Hop1.cell_type(1, i)
                            print(c)
                            if c == 0:
                                tk.messagebox.showinfo("error", "DataCompleteness-Hop1 Input sheet have empty values")
                                print(" empty cell")
                            else:
                                continue
                        except:
                            c = 0
                else:
                    tk.messagebox.showinfo("Important Message",
                                           "DataCompleteness-Hop1 Input sheet does not have any input to run the test. ")

    def data_cleansing_test(self):
        if self.cleansing_check.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)

            if self.cleansing_check.get() == 1:
                print("hop 2 is selected with variable " + str(self.integrity_check_1.get()))
                sheet_Hop2 = wb.sheet_by_name('Hop2_DataCleansing_Input')
                if sheet_Hop2.nrows >= 2:
                    print("proceed")
                    for i in range(sheet_Hop2.ncols - 2):
                        try:
                            c = sheet_Hop2.cell_type(1, i)
                            print(c)
                            if c == 0:
                                tk.messagebox.showinfo("error", "DataCleansing-Hop2 Input sheet have empty values")
                                print(" empty cell")
                            else:
                                continue
                        except:
                            c = 0
                else:
                    tk.messagebox.showinfo("Important Message",
                                           "Hop2_DataCleansing_Input-Hop1 Input sheet does not have any input to run the test. ")

    # def data_transformation_test(self):
    #     if self.transformation_check.get() == 0:
    #
    #         tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
    #     else:
    #         filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
    #         wb = xlrd.open_workbook(filename)
    #         print(len(wb.sheet_names()))
    #         print(wb.sheet_names())
    #         # sheet = wb.sheet_by_index(0)
    #         # print(sheet.nrows)
    #         # print("columns")
    #         # print(sheet.ncols)
    #         # os.system(filename)
    #
    #         if self.cleansing_check.get() == 1:
    #             print("hop 2 is selected with variable " + str(self.integrity_check_1.get()))
    #             sheet_Hop2 = wb.sheet_by_name('Hop2_DataCleansing_Input')
    #             if sheet_Hop2.nrows >= 2:
    #                 print("proceed")
    #                 for i in range(sheet_Hop2.ncols - 2):
    #                     try:
    #                         c = sheet_Hop2.cell_type(1, i)
    #                         print(c)
    #                         if c == 0:
    #                             tk.messagebox.showinfo("error", "DataCleansing-Hop2 Input sheet have empty values")
    #                             print(" empty cell")
    #                         else:
    #                             continue
    #                     except:
    #                         c = 0
    #             else:
    #                 tk.messagebox.showinfo("Important Message",
    #                                        "Hop2_DataCleansing_Input-Hop1 Input sheet does not have any input to run the test. ")
    #
    #
    #

    def hop2_xml_data_load_test(self):
        if self.xml_data_load_check.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)

            if self.xml_data_load_check.get() == 1:
                print("hop 2 is selected with variable " + str(self.integrity_check_1.get()))
                sheet_Hop2 = wb.sheet_by_name('Hop2 - XML Data Load Test_Input')
                if sheet_Hop2.nrows >= 2:
                    print("proceed")
                    for i in range(sheet_Hop2.ncols):
                        try:
                            c = sheet_Hop2.cell_type(1, i)
                            print(c)
                            if c == 0:
                                tk.messagebox.showinfo("error", "XML Data Load Test-Hop2 Input sheet have empty values")
                                print(" empty cell")
                            else:
                                continue
                        except:
                            c = 0
                else:
                    tk.messagebox.showinfo("Important Message",
                                           "Hop2 - XML Data Load Test Input sheet does not have any input to run the test. ")

    def hop3_stored_proc_test(self):
        if self.stored_proc_check.get() == 0:

            tk.messagebox.showinfo(" Error ", " Please Select the Applicable Test Phase for execution")
        else:
            filename = filedialog.askopenfilename(initialdir="C:/", title="select file")
            wb = xlrd.open_workbook(filename)
            print(len(wb.sheet_names()))
            print(wb.sheet_names())
            # sheet = wb.sheet_by_index(0)
            # print(sheet.nrows)
            # print("columns")
            # print(sheet.ncols)
            # os.system(filename)

            if self.stored_proc_check.get() == 1:
                print("hop 3 is selected with variable " + str(self.integrity_check_1.get()))
                sheet_Hop3 = wb.sheet_by_name('Hop3_Store_Proc_Input')
                if sheet_Hop3.nrows >= 2:
                    print("proceed")
                    for i in range(sheet_Hop3.ncols - 1):
                        try:
                            c = sheet_Hop3.cell_type(1, i)
                            print(c)
                            if c == 0:
                                tk.messagebox.showinfo("error", "Store_Proc-Hop3 Input sheet have empty values")
                                print(" empty cell")
                            else:
                                continue
                        except:
                            c = 0
                else:
                    tk.messagebox.showinfo("Important Message",
                                           "Store_Proc_Input-Hop3 Input sheet does not have any input to run the test. ")

    def display(self):
        self.frame_a.pack()
        self.frame_title.pack(anchor="nw")
        self.frame_b.pack(side="left", anchor="nw")
        self.framebutton.pack(side="right", anchor="ne")

        self.frameCheck.config(borderwidth=3.5, relief="ridge")
        self.frameCheck.pack(side="top", pady=11, anchor="w")
        self.frameCheck2.config(borderwidth=3.5, relief="ridge")
        self.frameCheck2.pack(pady=7, anchor="w")
        self.frameCheck3.config(borderwidth=3.5, relief="ridge")
        self.frameCheck3.pack(pady=7, anchor="w")
        self.frameCheck4.config(borderwidth=3.5, relief="ridge")
        self.frameCheck4.pack(pady=7, anchor="w")
        self.frameCheck5.config(borderwidth=3.5, relief="ridge")
        self.frameCheck5.pack(pady=7, anchor="w")
        self.frameCheck6.config(borderwidth=3.5, relief="ridge")
        self.frameCheck6.pack(pady=7, anchor="w")
        self.frameCheck7.config(borderwidth=3.5, relief="ridge")
        self.frameCheck7.pack(pady=7, anchor="w")
        self.frameCheck8.config(borderwidth=3.5, relief="ridge")
        self.frameCheck8.pack(pady=7, ipadx=20, anchor="w")

        self.window.geometry("1260x630")
        self.window.resizable(0, 0)
        self.window.mainloop()


if __name__ == '__main__':
    ui = ui_frame()
    ui.a()
    ui.b()

    ui.checkbox_frame1()
    ui.checkbox_frame2()
    ui.checkbox_frame3()
    ui.checkbox_frame4()
    ui.checkbox_frame5()
    ui.checkbox_frame6()
    ui.checkbox_frame7()
    ui.checkbox_frame8()

    ui.buttons_frame()

    ui.display()
