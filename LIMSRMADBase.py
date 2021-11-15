"""
LIMSRMADBase is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to act as an accessor to any and all related documentation regarding RMA evaluations that have been
performed in the Laboratory. It also has the capability of opening a webpage to the RMA software, Service Management.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import glob
import os
import os.path
import subprocess as sub

from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk


class AppRMADatabase:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # ==============================RMA DATABASE============================== #

    # Command to Open Window to Allow User to Perform RMA Based Searches
    def rma_database(self):
        global RMASearch, rma_summary_sheet_year, rma_evaluation_year, rma_number
        rma_summary_sheet_year = StringVar()
        rma_evaluation_year = StringVar()
        rma_number = StringVar()
        from LIMSHomeWindow import AppCommonCommands
        from LIMSHomeWindow import AppHomeWindow
        from LIMSHelpWindows import AppHelpWindows
        old_window = AppHomeWindow()
        old_window.home_window_hide()
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

    # ......................Main Window Properties............................ #

        RMASearch = Toplevel()
        RMASearch.title("RMA Evaluation File Search")
        RMASearch.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 345
        height = 380
        screen_width = RMASearch.winfo_screenwidth()
        screen_height = RMASearch.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        RMASearch.geometry("%dx%d+%d+%d" % (width, height, x, y))
        RMASearch.focus_force()
        RMASearch.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(RMASearch))

    # .........................Menu Bar Creation.................................. #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(RMASearch, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(RMASearch)),
                                           ("Logout", lambda: acc.software_signout(RMASearch)),
                                           ("Quit", lambda: acc.software_close(RMASearch))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.rma_database_help())])

    # .........................Frame Creation................................. #

        rma_search_frame = LabelFrame(RMASearch, text="Open RMA Summary Sheet by Year",
                                      relief=SOLID, bd=1, labelanchor="n")
        rma_search_frame.grid(row=0, column=0, rowspan=1, columnspan=1, pady=5)

        rma_search_frame2 = LabelFrame(RMASearch, text="Search for RMA Evaluations by " +
                                                       "Year \n or \n by Year and RMA Number",
                                       relief=SOLID, bd=1, labelanchor="n")
        rma_search_frame2.grid(row=6, column=0, rowspan=1, columnspan=1, pady=5, padx=8)

        rma_search_frame3 = LabelFrame(RMASearch, text="Connect to Service Management",
                                       relief=SOLID, bd=1, labelanchor="n")
        rma_search_frame3.grid(row=10, column=0, rowspan=1, columnspan=1, pady=5)

    # .....................Labels and Entries................................. #

        # Ask for Valid RMA Summary Sheet Year
        lbl_rma_summary_sheet_year = ttk.Label(rma_search_frame, text="RMA Year:", font=('arial', 12))
        lbl_rma_summary_sheet_year.grid(row=1, padx=8, pady=8)
        rma_summary_sheet_year_val = ttk.Entry(rma_search_frame, textvariable=rma_summary_sheet_year, font=14)
        rma_summary_sheet_year_val.grid(row=1, column=1)
        rma_summary_sheet_year_val.config(width=15)
        rma_summary_sheet_year_val.focus()

        # This label is a dummy label that does nothing but help format the window
        dummy = Label(rma_search_frame)
        dummy.grid(row=1, column=2)
        dummy.config(width=2)

        # Ask for Valid RMA Evaluation Year for Specific RMA
        lbl_rma_evaluation_year = ttk.Label(rma_search_frame2, text="RMA Year:", width=16, font=('arial', 12),
                                            anchor="n")
        lbl_rma_evaluation_year.grid(row=1, padx=8, pady=8)
        rma_evaluation_year_val = ttk.Entry(rma_search_frame2, textvariable=rma_evaluation_year, font=14)
        rma_evaluation_year_val.grid(row=1, column=1)
        rma_evaluation_year_val.config(width=15)

        # This label is a dummy label that does nothing but help format the window
        dummy3 = ttk.Label(rma_search_frame2)
        dummy3.grid(row=1, column=2)
        dummy3.config(width=1)
        dummy4 = ttk.Label(rma_search_frame2)
        dummy4.grid(row=1, column=3)
        dummy4.config(width=1)

        # Ask for Valid RMA Number for Specific RMA Evaluation
        lbl_rma_number = ttk.Label(rma_search_frame2, text="RMA Number:", width=16, font=('arial', 12), anchor="n")
        lbl_rma_number.grid(row=2)
        rma_number_val = ttk.Entry(rma_search_frame2, textvariable=rma_number, font=14)
        rma_number_val.grid(row=2, column=1)
        rma_number_val.config(width=15)
    
        # This label is a dummy label that does nothing but help format the window
        dummy5 = ttk.Label(rma_search_frame3)
        dummy5.grid(row=0, column=1, padx=8)
        dummy5.config(width=1)
        dummy6 = ttk.Label(rma_search_frame3)
        dummy6.grid(row=0, column=3)
        dummy6.config(width=1)

    # Add additional create RMA function including RMA form for all RMAs generated?

    # .............................Buttons................................... #

        btn_rma_summary_sheet_year = ttk.Button(rma_search_frame, text="Open", width=20,
                                                command=lambda: self.rma_summary_sheet_year_search())
        btn_rma_summary_sheet_year.bind("<Return>", lambda event: self.rma_summary_sheet_year_search())
        btn_rma_summary_sheet_year.grid(row=2, columnspan=3, pady=5)

        btn_rma_search = ttk.Button(rma_search_frame2, text="Search", width=20,
                                    command=lambda: self.rma_year_and_rma_number_search())
        btn_rma_search.bind("<Return>", lambda event: self.rma_year_and_rma_number_search())
        btn_rma_search.grid(row=3, column=0, columnspan=5, pady=5)
    
        btn_service_management_connect = ttk.Button(rma_search_frame3, text="Connect", width=20,
                                                    command=lambda: self.service_management_connection())
        btn_service_management_connect.bind("<Return>", lambda event: self.service_management_connection())
        btn_service_management_connect.grid(pady=5, row=0, column=2)
    
        btn_backout_rma = ttk.Button(RMASearch, text="Back", width=20, command=lambda: acc.return_home(RMASearch))
        btn_backout_rma.bind("<Return>", lambda event: acc.return_home(RMASearch))
        btn_backout_rma.grid(pady=5, row=12, column=0, columnspan=1)

    # ----------------------------------------------------------------------- #

    # Opens RMA Summary Sheet of Year Entered
    def rma_summary_sheet_year_search(self):

        if rma_summary_sheet_year.get().isdigit() is True and len(rma_summary_sheet_year.get()) == 4:
            f = glob.glob(os.path.join(r"\\BDC5\RMA Database\\" + rma_summary_sheet_year.get()+"\\",
                                       rma_summary_sheet_year.get() + " RMA summary.*"))[0]
            os.startfile(f)
        elif rma_summary_sheet_year.get().isdigit() is False and len(rma_summary_sheet_year.get()) == 4:
            tm.showerror("Invalid Year Entry", "Please enter a valid year that does not contain any letters")
        elif len(rma_summary_sheet_year.get()) < 4:
            tm.showerror("Invalid Year Entry", "Please enter a valid year for the RMA Summary Sheet search.")
        elif len(rma_summary_sheet_year.get()) > 4:
            tm.showerror("Too Many Characters", "Please enter a valid year for the RMA Summary Sheet search.")
        elif rma_summary_sheet_year.get().isupper() or rma_summary_sheet_year.get().islower():
            tm.showerror("Year Contains Letters", "Please enter a valid year for the RMA Summary Sheet search that \
does not contain letters.")

    # ----------------------------------------------------------------------- #

    # Opens RMA Evaluation File Folder Input By User
    def rma_year_and_rma_number_search(self):

        if rma_evaluation_year.get() != "" and rma_evaluation_year.get().isdigit() is True and \
                len(rma_evaluation_year.get()) == 4 and rma_number.get() != "" and rma_number.get().isdigit() is True:
            f = glob.glob(os.path.join(r"\\BDC5\RMA Database\\", rma_evaluation_year.get()))[0]
            rma_file = glob.glob(os.path.join(f, rma_number.get()))[0]
            os.startfile(rma_file)
        elif rma_evaluation_year.get() != "" and rma_evaluation_year.get().isdigit() is True and rma_number.get() == "":
            f = glob.glob(os.path.join(r"\\BDC5\RMA Database\\", rma_evaluation_year.get()))[0]
            rma_file = glob.glob(os.path.join(f, rma_number.get()))[0]
            os.startfile(rma_file)
        elif rma_evaluation_year.get() == "" and rma_number.get() == "":
            f = glob.glob(os.path.join(r"\\BDC5\RMA Database\\", rma_evaluation_year.get()))[0]
            rma_file = glob.glob(os.path.join(f, rma_number.get()))[0]
            os.startfile(rma_file)
        elif len(rma_evaluation_year.get()) < 4:
            tm.showerror("Invalid Year Entry", "Please enter a valid year for the RMA you are looking for.")
        elif len(rma_evaluation_year.get()) > 4:
            tm.showerror("Too Many Characters", "Please enter a valid year for the RMA you are looking for.")
        elif rma_evaluation_year.get() == "" and rma_number.get() != "" and rma_number.get().isdigit() is True:
            tm.showerror("RMA Year Required", "Please provide a year for the RMA you are looking for.")
        elif rma_evaluation_year.get().isdigit() is False or rma_number.get().isdigit() is False:
            tm.showerror("Invalid Entry(s)", "Please enter a valid year and/or number that does not contain letters.")

    # ----------------------------------------------------------------------- #
 
    # Connect to Service Management
    def service_management_connection(self):

        sub.Popen(r'"C:\\Program Files\\Internet Explorer\\iexplore.exe"http://ibmsoftsol/sc/login.asp?DomainId=DWYER')
