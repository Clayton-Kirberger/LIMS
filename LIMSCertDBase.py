"""
LIMSCertDBase is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to act as an assessor to any and all information we have on our BDC5 network that is in regards to
the certificates of calibration generated by the Engineering Laboratory.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import glob
import os
import os.path
import subprocess as sub
from datetime import datetime

import LIMSVarConfig
import win32com.client
from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk


class AppCertificateDatabase:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # ==========================CERTIFICATE DATABASE========================== #

    # Certificate of Calibration Query Window & Process
    def certificate_database(self):
        global CertSearch, certificate_number_val, certificate_year_val, certificate_year_log_val
        certificate_year = StringVar()
        certificate_year_log = StringVar()
        certificate_number = StringVar()
        from LIMSHomeWindow import AppCommonCommands
        from LIMSHomeWindow import AppHomeWindow
        from LIMSHelpWindows import AppHelpWindows
        old_window = AppHomeWindow()
        old_window.home_window_hide()
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

    # ......................Main Window Properties............................ #

        CertSearch = Toplevel()
        CertSearch.title("Certificate of Calibration File Search")
        CertSearch.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 325
        height = 350
        screen_width = CertSearch.winfo_screenwidth()
        screen_height = CertSearch.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        CertSearch.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CertSearch.focus_force()
        CertSearch.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CertSearch))

    # .........................Menu Bar Creation.................................. #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(CertSearch, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CertSearch)),
                                           ("Logout", lambda: acc.software_signout(CertSearch)),
                                           ("Quit", lambda: acc.software_close(CertSearch))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.certificate_database_help())])

    # .........................Frame Creation................................. #

        certificate_search_frame = LabelFrame(CertSearch, text="Search for Certificates of Calibration by Year",
                                              relief=SOLID, bd=1, labelanchor="n")
        certificate_search_frame.grid(row=0, column=0, rowspan=2, columnspan=3, padx=7, pady=6)

        certificate_search_frame_1 = LabelFrame(CertSearch, text="Open Log of Certificates by Year", relief=SOLID,
                                                bd=1, labelanchor="n")
        certificate_search_frame_1.grid(row=3, column=0, rowspan=2, columnspan=3, padx=7, pady=6)

        certificate_search_frame_2 = LabelFrame(CertSearch, text="Search for Certificates by Certificate Number",
                                                relief=SOLID, bd=1, labelanchor="n")
        certificate_search_frame_2.grid(row=6, column=0, rowspan=2, columnspan=3, padx=7, pady=6)

    # ................Labels and Entries for Current Window................... #

        # Ask for folder Year that Certificates of Calibration were Issued
        lbl_certificate_year = ttk.Label(certificate_search_frame, text="Certificate Year:", font=('arial', 12))
        lbl_certificate_year.grid(row=1, padx=5, pady=5)
        certificate_year_val = ttk.Entry(certificate_search_frame, textvariable=certificate_year, font=14)
        certificate_year_val.grid(row=1, column=1)
        certificate_year_val.config(width=15)
        certificate_year_val.focus()

        # This label is a dummy label that does nothing but help format the window
        dummy = ttk.Label(certificate_search_frame)
        dummy.grid(row=1, column=2)
        dummy.config(width=2)

        # Ask for Year of Log for Certificates of Calibration
        lbl_certificate_year_log = ttk.Label(certificate_search_frame_1, text="Certificate Year:", font=('arial', 12))
        lbl_certificate_year_log.grid(row=4, padx=5, pady=5)
        certificate_year_log_val = ttk.Entry(certificate_search_frame_1, textvariable=certificate_year_log, font=14)
        certificate_year_log_val.grid(row=4, column=1)
        certificate_year_log_val.config(width=15)

        # This label is a dummy label that does nothing but help format the window
        dummy1 = ttk.Label(certificate_search_frame_1)
        dummy1.grid(row=4, column=2)
        dummy1.config(width=2)

        # Ask for Certificate of Calibration Number
        lbl_certificate_number = ttk.Label(certificate_search_frame_2, text="Certificate Number:", font=('arial', 12))
        lbl_certificate_number.grid(row=7, padx=5, pady=5)
        certificate_number_val = ttk.Entry(certificate_search_frame_2, textvariable=certificate_number, font=14)
        certificate_number_val.bind("<KeyRelease>", lambda event: acc.all_caps(certificate_number))
        certificate_number_val.grid(row=7, column=1)
        certificate_number_val.config(width=15)

        # These labels are dummy labels that do nothing but help format the window
        dummy2 = ttk.Label(certificate_search_frame_2)
        dummy2.grid(row=7, column=2)
        dummy2.config(width=2)
        dummy3 = ttk.Label(CertSearch)
        dummy3.grid(row=9)
        dummy3.config(width=2)
        dummy4 = ttk.Label(CertSearch)
        dummy4.grid(row=9, column=2)
        dummy4.config(width=2)

    # ................Button with Functions for this Window................... #

        btn_certificate_year_search = ttk.Button(certificate_search_frame, text="Search", width=20,
                                                 command=lambda: self.certificate_year_search())
        btn_certificate_year_search.bind("<Return>", lambda event: self.certificate_year_search())
        btn_certificate_year_search.grid(row=2, columnspan=3, pady=5)

        btn_certificate_year_log_search = ttk.Button(certificate_search_frame_1, text="Open", width=20,
                                                     command=lambda: self.certificate_year_log_search())
        btn_certificate_year_log_search.bind("<Return>", lambda event: self.certificate_year_log_search())
        btn_certificate_year_log_search.grid(row=5, columnspan=3, pady=5)

        btn_certificate_number_search = ttk.Button(certificate_search_frame_2, text="Open", width=20,
                                                   command=lambda: self.certificate_number_search())
        btn_certificate_number_search.bind("<Return>", lambda event: self.certificate_number_search())
        btn_certificate_number_search.grid(row=8, columnspan=3, pady=5)

        btn_backout_certificate_database = ttk.Button(CertSearch, text="Back to Main Menu", width=20,
                                                      command=lambda: acc.return_home(CertSearch))
        btn_backout_certificate_database.bind("<Return>", lambda event: acc.return_home(CertSearch))
        btn_backout_certificate_database.grid(pady=5, padx=5, row=9, column=1, columnspan=1)

    # ----------------------------------------------------------------------- #

    # Opens folder for Year of Calibration Certificates Input By User
    def certificate_year_search(self):

        certificate_year = certificate_year_val.get()
        if certificate_year.isdigit() and len(certificate_year) == 4:
            sub.Popen(r'explorer /open, "\\\\BDC5\certdbase\"' + certificate_year_val.get() + "")
        else:
            tm.showerror("Invalid Entry", "Please provide a valid year of calibration for desired directory.")

    # ----------------------------------------------------------------------- #

    # Opens records sheet for Year of Calibration Certificate Input By User
    def certificate_year_log_search(self):

        if certificate_year_log_val.get().isdigit() and len(certificate_year_log_val.get()) == 4:
            f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\\" + certificate_year_log_val.get() + "\\",
                                       certificate_year_log_val.get() + " Certificates of calibration.*"))[0]
            os.startfile(f)
        else:
            tm.showerror("Invalid Entry", "Please provide a valid year of calibration for certificate database.")

    # ----------------------------------------------------------------------- #

    # Opens Certificate of Calibration Based on Year and User Input
    def certificate_number_search(self):

        if "22DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2022\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "21DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2021\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "20DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2020\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "19DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2019\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "18DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2018\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "17DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2017\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "16DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2016\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "15DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2015\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "14DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2014\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "13DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2013\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "12DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2012\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "11DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2011\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "10DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2010\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "09DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2009\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "08DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2008\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "07DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2007\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "06DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2006\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "05DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2005\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "04DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2004\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "03DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2003\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "02DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2002\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "01DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2001\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        elif "00DWY00" in certificate_number_val.get():
            try:
                f = glob.glob(os.path.join(r"\\\\BDC5\certdbase\2000\\", certificate_number_val.get() + ".*"))[0]
                os.startfile(f)
            except IndexError:
                self.invalid_entry_message()

        else:
            self.invalid_entry_message()

    # ----------------------------------------------------------------------- #

    # Error Message Shown for Invalid Entry
    def invalid_entry_message(self):
        tm.showerror("Invalid Entry", "Please provide a valid certificate number for \
calibration performed and documented in Michigan City, IN.")

    # ----------------------------------------------------------------------- #

    # Command to Search for Next Available Cert Number
    def certificate_number_helper(self):

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        year = datetime.today().year
        wkbook = excel.Workbooks.Open('\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls' %(year, year))
        sheet = wkbook.Sheets("Certifications")


        for row in range(1, 3000):
            cell = sheet.Cells(row, 5)
            if cell.Value is None:
                col = 5  # cell.column
                cell = sheet.Cells(row, col - 4)
                LIMSVarConfig.certificate_of_calibration_number = str(cell)
                sheet.Cells(row, 5).Value = LIMSVarConfig.external_customer_name
                sheet.Cells(row, 6).Value = LIMSVarConfig.external_customer_sales_order_number_helper
                sheet.Cells(row, 7).Value = LIMSVarConfig.external_customer_rma_number
                sheet.Cells(row, 9).Value = LIMSVarConfig.calibration_date_helper
                sheet.Cells(row, 12).Value = "In Progress"
                break
        wkbook.Close(True)

    # ----------------------------------------------------------------------- #

    # Command to Search for Next Available Cert Number
    def certificate_serial_number_helper(self):

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        year = datetime.today().year
        wkbook = excel.Workbooks.Open('\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls' %(year, year))
        sheet = wkbook.Sheets("Certifications")

        # Initialize Location of Serial Number String and Serial Number Values
        initial_serial_number_string = []
        serial_number_string = []
        serial_number_values = []

        # Loop until new cert number is found
        i = 1
        cell = sheet.Cells(i, 5)
        for i in range(1, 3000):
            if cell.Value is None:
                x = i  # cell.row
                y = 5  # cell.column
                cell = sheet.Cells(x, y - 4)
                LIMSVarConfig.certificate_of_calibration_number = str(cell)
                sheet.Cells(x, 5).Value = LIMSVarConfig.external_customer_name
                sheet.Cells(x, 6).Value = LIMSVarConfig.external_customer_sales_order_number_helper
                sheet.Cells(x, 7).Value = LIMSVarConfig.external_customer_rma_number
                sheet.Cells(x, 9).Value = LIMSVarConfig.calibration_date_helper
                sheet.Cells(x, 12).Value = "In Progress"
                break
            else:
                i += 1
                cell = sheet.Cells(i, 5)

        # Log all existing serial numbers to array
        i = 1
        cell = sheet.Cells(i, 2)
        for i in range(1, 3000):
            if cell.Value is not None:
                initial_serial_number_string.append(str(cell).strip())
                i += 1
                cell = sheet.Cells(i, 2)

        # Parse through list to only find serial numbers starting with M0
        for entry in initial_serial_number_string:
            if str("M0") in entry:
                serial_number_string.append(str(entry).replace("M0", "").replace("M", ""))

        # Log serial number values to array
        for entry in serial_number_string:
            if int(5) >= len(entry) > int(4):
                serial_number_values.append(entry)

        # Provide new serial number
        new_serial_number = (int(max(serial_number_values)) + 1)
        LIMSVarConfig.device_serial_number = ("M0" + str(new_serial_number))

        i = 1
        cell = sheet.Cells(i, 1)
        for i in range(1, 3000):
            if cell.Value == LIMSVarConfig.certificate_of_calibration_number:
                x = i
                sheet.Cells(x, 2).Value = LIMSVarConfig.device_serial_number
                break
            else:
                i += 1
                cell = sheet.Cells(i, 1)
    
        wkbook.Close(True)

    # ----------------------------------------------------------------------- #

    # Command to Search for Exisiting Certificates
    def certificate_number_checker(self):

        excel = win32com.client.Dispatch("Excel.Application")
        year = datetime.today().year
        wkbook = excel.Workbooks.Open('\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls' %(year, year))
        sheet = wkbook.Sheets("Certifications")
        print('Cert Database opened')

        for i in range(1, 3000):
            if (sheet.Cells(i, 5).Value == LIMSVarConfig.external_customer_name and sheet.Cells(i, 6).Value ==
                    LIMSVarConfig.external_customer_sales_order_number_helper and
                    sheet.Cells(i, 12).Value == "In Progress"):
                cell = sheet.Cells(i, 1)
                LIMSVarConfig.certificate_of_calibration_number = str(cell)
                break
            else:
                i += 1
        wkbook.Close(True)

    # ----------------------------------------------------------------------- #

    # Command to Search for Exisiting Certificates and Serial Numbers
    def certificate_serial_number_checker(self):

        excel = win32com.client.Dispatch("Excel.Application")
        year = datetime.today().year
        wkbook = excel.Workbooks.Open('\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls' %(year, year))
        sheet = wkbook.Sheets("Certifications")

        for i in range(1, 3000):
            if (sheet.Cells(i, 5).Value == LIMSVarConfig.external_customer_name and sheet.Cells(i, 6).Value ==
                    LIMSVarConfig.external_customer_sales_order_number_helper and
                    sheet.Cells(i, 12).Value == "In Progress"):
                cell = sheet.Cells(i, 1)
                LIMSVarConfig.certificate_of_calibration_number = str(cell)
                x = i
                LIMSVarConfig.device_serial_number = str(sheet.Cells(x, 2).Value)
                break
            else:
                i += 1
        wkbook.Close(True)

    # ----------------------------------------------------------------------- #

    # Command to search for existing calibration and provide completed certificate information
    def certificate_number_fill_in(self):

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        year = datetime.today().year
        wkbook = excel.Workbooks.Open('\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls' %(year, year))
        sheet = wkbook.Sheets("Certifications")

        i = 1
        cell = sheet.Cells(i, 1)
        for i in range(1, 3000):
            if cell.Value == LIMSVarConfig.certificate_of_calibration_number:
                x = i
                sheet.Cells(x, 2).Value = LIMSVarConfig.device_identification_number_helper
                sheet.Cells(x, 3).Value = LIMSVarConfig.device_date_code_helper
                sheet.Cells(x, 4).Value = LIMSVarConfig.device_model_number_helper
                sheet.Cells(x, 8).Value = LIMSVarConfig.device_date_received_helper
                sheet.Cells(x, 9).Value = LIMSVarConfig.calibration_date_helper
                sheet.Cells(x, 10).Value = LIMSVarConfig.certificate_technician_name
                sheet.Cells(x, 11).Value = LIMSVarConfig.calibration_time_helper
                sheet.Cells(x, 12).Value = ""
                break
            else:
                i += 1
                cell = sheet.Cells(i, 1)
            
        wkbook.Close(True)

    # ----------------------------------------------------------------------- #

    # Command to fill in "Failed" field in certdbase
    def certificate_failure(self):

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        year = datetime.today().year
        wkbook = excel.Workbooks.Open('\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls' %(year, year))
        sheet = wkbook.Sheets("Certifications")

        i = 1
        cell = sheet.Cells(i, 1)
        for i in range(1, 3000):
            if cell.Value == LIMSVarConfig.certificate_of_calibration_number:
                x = i
                sheet.Cells(x, 12).Value = "FAIL"
                break
            else:
                i += 1
                cell = sheet.Cells(i, 1)
            
        wkbook.Close(True)
