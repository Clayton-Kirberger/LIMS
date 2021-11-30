"""
LIMSCertCreation is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering
LIMS. With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled
at will to add any necessary additions that are deemed to fit the requirements of the laboratory. This module serves
as the meat and potatoes of the Calibration GUI for all technicians in the laboratory. This function contains other
submodules that are useful tools for the personnel generating certificates of calibration.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import datetime
import os
import os.path
from datetime import datetime

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppCalibrationModule:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =====================CERTIFICATE OF CALIBRATION CREATION======================== #

    # Calibration Type Selection
    def customer_calibration_type_selection(self, window):
        window.withdraw()

        global CalTypeSel, customer_calibration_selection

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        CalTypeSel = Toplevel()
        CalTypeSel.title("Customer Selection")
        CalTypeSel.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 290
        height = 135
        screen_width = CalTypeSel.winfo_screenwidth()
        screen_height = CalTypeSel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        CalTypeSel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CalTypeSel.focus_force()
        CalTypeSel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CalTypeSel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(CalTypeSel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CalTypeSel)),
                                           ("Logout", lambda: acc.software_signout(CalTypeSel)),
                                           ("Quit", lambda: acc.software_close(CalTypeSel))])

        # ......................Frame Creation...................... #

        customer_selection_frame = LabelFrame(CalTypeSel, text="Select Customer Type from Drop Down List",
                                              relief=SOLID, bd=1, labelanchor="n")
        customer_selection_frame.grid(row=0, column=0, rowspan=2, columnspan=2, padx=10, pady=5)

        # .......................Drop Down Lists................................. #

        customer_calibration_selection = ttk.Combobox(customer_selection_frame, values=[" ", "Customer", "Internal"])
        acc.always_active_style(customer_calibration_selection)
        customer_calibration_selection.configure(state="active", width=15)
        customer_calibration_selection.focus()
        customer_calibration_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(customer_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_customer_selection = ttk.Button(customer_selection_frame, text="Open", width=15,
                                                 command=lambda: self.open_customer_calibration_option())
        btn_open_customer_selection.bind("<Return>", lambda event: self.open_customer_calibration_option())
        btn_open_customer_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_customer_selection = ttk.Button(customer_selection_frame, text="Back to Main Menu", width=20,
                                                    command=lambda: acc.return_home(CalTypeSel))
        btn_backout_customer_selection.bind("<Return>", lambda event: acc.return_home(CalTypeSel))
        btn_backout_customer_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle customer calibration type selection
    def open_customer_calibration_option(self):
        if customer_calibration_selection.get() == "Customer":
            self.sales_order_import_option_window(CalTypeSel)
        elif customer_calibration_selection.get() == "Internal":
            from LIMSCCCreation import AppCCModule
            accm = AppCCModule()
            accm.internal_customer_selection(CalTypeSel)
        else:
            tm.showerror("No Calibration Selection", "Please select the type of customer that you will be performing \
calibration work for.")

    # ----------------------------------------------------------------------------- #

    # Sales Order Information Query Process
    def sales_order_import_option_window(self, window):
        window.withdraw()

        global SOSearch, as400_user_val, as400_pass_val, sales_order_val, sales_order

        as400_username = StringVar()
        as400_password = StringVar()
        sales_order = StringVar()

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        LIMSVarConfig.clear_external_customer_sales_order_variables()
        LIMSVarConfig.customer_selection_check = int(0)

        # ........................Main Window Properties......................... #

        SOSearch = Toplevel()
        SOSearch.title("Customer Sales Order Information Search")
        SOSearch.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 355
        height = 220
        screen_width = SOSearch.winfo_screenwidth()
        screen_height = SOSearch.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        SOSearch.geometry("%dx%d+%d+%d" % (width, height, x, y))
        SOSearch.focus_force()
        SOSearch.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(SOSearch))

        # .........................Menu Bar Creation.................................. #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(SOSearch, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(SOSearch)),
                                           ("Logout", lambda: acc.software_signout(SOSearch)),
                                           ("Quit", lambda: acc.software_close(SOSearch))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.sales_order_import_option_help())])

        # ..........................Frame Creation............................... #

        sales_order_search_frame = LabelFrame(SOSearch, text="Log New Customer Information from AS400",
                                              relief=SOLID, bd=1, labelanchor="n")
        sales_order_search_frame.grid(row=0, column=0, rowspan=6, columnspan=4, padx=8, pady=5)

        # ........................Labels and Entries............................. #

        # Ask for AS400 Username Information
        lbl_as400_user = ttk.Label(sales_order_search_frame, text="AS400 Username:", font=('arial', 12))
        lbl_as400_user.grid(row=1, padx=5, pady=5)
        as400_user_val = ttk.Entry(sales_order_search_frame, textvariable=as400_username, font=14)
        as400_user_val.grid(row=1, column=1, columnspan=2)
        as400_user_val.config(width=15)
        as400_user_val.focus()

        # Ask for AS400 Password Information
        lbl_as400_pass = ttk.Label(sales_order_search_frame, text="AS400 Password:", font=('arial', 12))
        lbl_as400_pass.grid(row=2, padx=5, pady=5)
        as400_pass_val = ttk.Entry(sales_order_search_frame, textvariable=as400_password, show="*", font=14)
        as400_pass_val.grid(row=2, column=1, columnspan=2)
        as400_pass_val.config(width=15)

        # Ask for Valid Sales Order Number
        lbl_as400_sales_order = ttk.Label(sales_order_search_frame, text="Customer Sales Order:", font=('arial', 12))
        lbl_as400_sales_order.grid(row=3, padx=5, pady=5)
        sales_order_val = ttk.Entry(sales_order_search_frame, textvariable=sales_order, font=14)
        sales_order_val.bind("<KeyRelease>", lambda event: acc.all_caps(sales_order))
        sales_order_val.grid(row=3, column=1, columnspan=2)
        sales_order_val.config(width=15)

        # This label is a dummy label that does nothing but help format the window
        dummy = Label(sales_order_search_frame)
        dummy.grid(row=1, column=3)
        dummy.config(width=2)

        # ...............................Buttons.................................. #

        btn_backout = ttk.Button(sales_order_search_frame, text="Back",
                                 command=lambda: self.customer_calibration_type_selection(SOSearch))
        btn_backout.bind("<Return>", lambda event: self.customer_calibration_type_selection(SOSearch))
        btn_backout.grid(row=5, column=0, columnspan=1, pady=5)
        btn_backout.config(width=15)

        btn_next = ttk.Button(sales_order_search_frame, text="Search AS400", command=lambda: self.as400_executable())
        btn_next.bind("<Return>", lambda event: self.as400_executable())
        btn_next.grid(row=5, column=1, columnspan=1, pady=5)
        btn_next.config(width=15)

        btn_external_customer_information = ttk.Button(sales_order_search_frame, text="Create Customer Certificate",
                                                       command=lambda: self.external_customer_information(SOSearch))
        btn_external_customer_information.bind("<Return>", lambda event: self.external_customer_information(SOSearch))
        btn_external_customer_information.grid(row=6, column=0, columnspan=3, pady=5)
        btn_external_customer_information.config(width=25)

    # ----------------------------------------------------------------------- #

    # Open AS400 and Extract Information to .txt File
    def as400_executable(self):

        LIMSVarConfig.as400_username_helper = as400_user_val.get()
        LIMSVarConfig.as400_password_helper = as400_pass_val.get()
        LIMSVarConfig.as400_sales_order_helper = sales_order_val.get()

        if os.path.isfile("Sales Order Text Files\\" + sales_order_val.get() + ".txt") \
                is True:
            tm.showinfo("Customer Sales Order Information Search", "This sales order and its information has already \
been exported to a text file. Please continue to the next step.")
        elif LIMSVarConfig.as400_sales_order_helper == "" or LIMSVarConfig.as400_sales_order_helper == " ":
            tm.showerror("Customer Sales Order Information Search",
                         "Enter a customer sales order to initiate export process.")
        elif LIMSVarConfig.as400_sales_order_helper[0].isupper() is False:
            tm.showerror("Customer Sales Order Information Search", "Please capitalize the 'S' in the customer sales \
order, and then press the 'Search AS400' button again to initiate the export process")
        elif len(LIMSVarConfig.as400_sales_order_helper) < 7:
            tm.showerror("Customer Sales Order Information Search", "Please enter a valid customer Sales Order Number.")
        else:
            from LIMSAS400Search import AppAS400ExecutableHelper
            execute_command = AppAS400ExecutableHelper()
            execute_command.as400_executable_helper()

    # ----------------------------------------------------------------------- #

    # Allows User to Create New Customer Certificate By Importing Information
    def external_customer_information(self, window):
        window.withdraw()

        global CalCert, external_customer_sales_order_number_value, lbl_customer_name_value, \
            lbl_customer_address_value, lbl_customer_address_value_1, lbl_customer_address_value_2, \
            lbl_customer_city_value, lbl_state_country_zip_value, lbl_customer_purchase_order_value, \
            lbl_customer_rma_value, lbl_header_notes, lbl_header_notes_1, lbl_header_notes_2, lbl_header_notes_3, \
            lbl_header_notes_4, lbl_header_notes_5, lbl_item_notes, lbl_item_notes_1, lbl_item_notes_2, \
            lbl_item_notes_3, lbl_item_notes_4, lbl_item_notes_5, lbl_footer_notes, lbl_footer_notes_1, \
            lbl_footer_notes_2, lbl_footer_notes_3, lbl_footer_notes_4, lbl_footer_notes_5, lbl_pm_item_notes_value, \
            lbl_pm_item_notes_value_1, lbl_pm_item_notes_value_2, lbl_pm_item_notes_value_3, \
            lbl_pm_item_notes_value_4, lbl_pm_item_notes_value_5

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.......................... #

        CalCert = Toplevel()
        CalCert.title("Certificate of Calibration Creation")
        CalCert.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 960
        height = 515
        screen_width = CalCert.winfo_screenwidth()
        screen_height = CalCert.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        CalCert.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CalCert.focus_force()
        CalCert.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CalCert))

        # .........................Menu Bar Creation.................................. #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(CalCert, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CalCert)),
                                           ("Sales Order Search",
                                            lambda: self.sales_order_import_option_window(CalCert)),
                                           ("Logout", lambda: acc.software_signout(CalCert)),
                                           ("Quit", lambda: acc.software_close(CalCert))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.external_customer_certificate_information_help())])

        # ...........................Frame Creation............................... #

        customer_information_frame = LabelFrame(CalCert, text="Customer Information", relief=SOLID, bd=1,
                                                labelanchor="n")
        customer_information_frame.grid(row=0, column=0, rowspan=6, columnspan=6, padx=7, pady=5)

        customer_notes_frame = LabelFrame(CalCert, text="Customer Notes", relief=SOLID, bd=1, labelanchor="n")
        customer_notes_frame.grid(row=7, column=0, rowspan=9, columnspan=6, padx=7, pady=5)

        # .........................Labels and Entries............................. #

        # Input Sales Order Number of previously saved Sales Order information acquired from the AS400 Sales Order
        # query performed
        lbl_external_customer_sales_order_number = ttk.Label(customer_information_frame, text="Sales Order Number:",
                                                             font=('arial', 12))
        lbl_external_customer_sales_order_number.grid(row=1, padx=5)

        if sales_order_val.get() != "":
            customer_sales_order_number = sales_order
        else:
            customer_sales_order_number = StringVar()

        external_customer_sales_order_number_value = ttk.Entry(customer_information_frame,
                                                               textvariable=customer_sales_order_number,
                                                               font=('arial', 12))
        external_customer_sales_order_number_value.bind("<KeyRelease>",
                                                        lambda event: acc.all_caps(customer_sales_order_number))
        external_customer_sales_order_number_value.grid(pady=3, row=1, column=1)
        external_customer_sales_order_number_value.config(width=15)
        external_customer_sales_order_number_value.focus()

        # /////////////////////////////////////////////////////////////////////// #
        # The following is information provided by import (assuming a valid
        # sales order has been provided prior to clicking the "Import
        # Customer Information" button and that a file containing customer
        # information was actually created and saved in the proper directory)
        # /////////////////////////////////////////////////////////////////////// #

        # Customer Name
        lbl_customer_name = ttk.Label(customer_information_frame, text="Customer Name:", font=('arial', 12))
        lbl_customer_name.grid(row=2)
        lbl_customer_name_value = ttk.Label(customer_information_frame, text=LIMSVarConfig.external_customer_name,
                                            font=('arial', 12), anchor="n")
        lbl_customer_name_value.grid(row=2, column=1)
        lbl_customer_name_value.config(width=35)

        # Customer Address
        lbl_customer_address = ttk.Label(customer_information_frame, text="Customer Address:", font=('arial', 12))
        lbl_customer_address.grid(row=3)
        lbl_customer_address_value = ttk.Label(customer_information_frame, text=LIMSVarConfig.external_customer_address,
                                               font=('arial', 12))
        lbl_customer_address_value.grid(row=3, column=1)

        # Customer Address (cont.)
        lbl_customer_address_value_1 = ttk.Label(customer_information_frame,
                                                 text=LIMSVarConfig.external_customer_address_1, font=('arial', 12))
        lbl_customer_address_value_1.grid(pady=3, row=4, column=1)

        # Customer Address (cont.)
        lbl_customer_address_value_2 = ttk.Label(customer_information_frame,
                                                 text=LIMSVarConfig.external_customer_address_2, font=('arial', 12))
        lbl_customer_address_value_2.grid(pady=3, row=5, column=1)

        # Customer City
        lbl_customer_city = ttk.Label(customer_information_frame, text="Customer City:", font=('arial', 12))
        lbl_customer_city.grid(row=6)
        lbl_customer_city_value = ttk.Label(customer_information_frame,
                                            text=LIMSVarConfig.external_customer_city, font=('arial', 12))
        lbl_customer_city_value.grid(row=6, column=1)

        # Customer State, Country and Zip Code
        lbl_state_country_zip = ttk.Label(customer_information_frame, text="State/Country/Zip Code:",
                                          font=('arial', 12))
        lbl_state_country_zip.grid(row=6, column=2)
        lbl_state_country_zip_value = ttk.Label(customer_information_frame,
                                                text=LIMSVarConfig.external_customer_state_country_zip,
                                                font=('arial', 12))
        lbl_state_country_zip_value.grid(pady=3, row=6, column=3)

        # Customer P.O. Number
        lbl_customer_purchase_order = ttk.Label(customer_information_frame, text="Customer P.O.:", font=('arial', 12))
        lbl_customer_purchase_order.grid(row=2, column=2)
        lbl_customer_purchase_order_value = ttk.Label(customer_information_frame,
                                                      text=LIMSVarConfig.external_customer_po, font=('arial', 12))
        lbl_customer_purchase_order_value.grid(pady=3, row=2, column=3)

        # Customer RMA
        lbl_customer_rma = ttk.Label(customer_information_frame, text="RMA Number:", font=('arial', 12))
        lbl_customer_rma.grid(row=3, column=2)
        lbl_customer_rma_value = ttk.Label(customer_information_frame, text=LIMSVarConfig.external_customer_rma_number,
                                           font=('arial', 12), anchor="n")
        lbl_customer_rma_value.grid(pady=3, row=3, column=3)
        lbl_customer_rma_value.config(width=30)

        # Customer AS400 Header Notes (only available to technician)
        lbl_header_notes_title = ttk.Label(customer_notes_frame, text="Header Notes:", font=('arial', 12))
        lbl_header_notes_title.grid(pady=3, row=1, padx=5)
        lbl_header_notes = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_header_notes,
                                     font=('arial', 12))
        lbl_header_notes.grid(row=1, column=1)
        lbl_header_notes.config(width=30)

        lbl_header_notes_1 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_header_notes_1,
                                       font=('arial', 12))
        lbl_header_notes_1.grid(row=1, column=2)
        lbl_header_notes_1.config(width=30)

        lbl_header_notes_2 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_header_notes_2,
                                       font=('arial', 12))
        lbl_header_notes_2.grid(row=1, column=3)
        lbl_header_notes_2.config(width=30)

        lbl_header_notes_3 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_header_notes_3,
                                       font=('arial', 12))
        lbl_header_notes_3.grid(pady=3, row=2, column=1)
        lbl_header_notes_3.config(width=30)

        lbl_header_notes_4 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_header_notes_4,
                                       font=('arial', 12))
        lbl_header_notes_4.grid(row=2, column=2)
        lbl_header_notes_4.config(width=30)

        lbl_header_notes_5 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_header_notes_5,
                                       font=('arial', 12))
        lbl_header_notes_5.grid(row=2, column=3)
        lbl_header_notes_5.config(width=30)

        # Customer AS400 Item Notes (only available to technician)
        lbl_item_notes_title = ttk.Label(customer_notes_frame, text="Item Notes:", font=('arial', 12))
        lbl_item_notes_title.grid(pady=3, row=3)
        lbl_item_notes = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_item_notes,
                                   font=('arial', 12))
        lbl_item_notes.grid(row=3, column=1)

        lbl_item_notes_1 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_item_notes_1,
                                     font=('arial', 12))
        lbl_item_notes_1.grid(row=3, column=2)

        lbl_item_notes_2 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_item_notes_2,
                                     font=('arial', 12))
        lbl_item_notes_2.grid(row=3, column=3)

        lbl_item_notes_3 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_item_notes_3,
                                     font=('arial', 12))
        lbl_item_notes_3.grid(pady=3, row=4, column=1)

        lbl_item_notes_4 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_item_notes_4,
                                     font=('arial', 12))
        lbl_item_notes_4.grid(row=4, column=2)

        lbl_item_notes_5 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_item_notes_5,
                                     font=('arial', 12))
        lbl_item_notes_5.grid(row=4, column=3)

        # Customer AS400 Footer Notes (only available to technician)
        lbl_footer_notes_title = ttk.Label(customer_notes_frame, text="Footer Notes:", font=('arial', 12))
        lbl_footer_notes_title.grid(pady=3, row=5)
        lbl_footer_notes = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_footer_notes,
                                     font=('arial', 12))
        lbl_footer_notes.grid(row=5, column=1)

        lbl_footer_notes_1 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_footer_notes_1,
                                       font=('arial', 12))
        lbl_footer_notes_1.grid(row=5, column=2)

        lbl_footer_notes_2 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_footer_notes_2,
                                       font=('arial', 12))
        lbl_footer_notes_2.grid(row=5, column=3)

        lbl_footer_notes_3 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_footer_notes_3,
                                       font=('arial', 12))
        lbl_footer_notes_3.grid(pady=3, row=6, column=1)

        lbl_footer_notes_4 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_footer_notes_4,
                                       font=('arial', 12))
        lbl_footer_notes_4.grid(row=6, column=2)

        lbl_footer_notes_5 = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_footer_notes_5,
                                       font=('arial', 12))
        lbl_footer_notes_5.grid(row=6, column=3)

        # Customer AS400 PM Item Notes (only available to technician)
        lbl_pm_it_notes_title = ttk.Label(customer_notes_frame, text="PM Item Notes:", font=('arial', 12))
        lbl_pm_it_notes_title.grid(pady=3, row=7, padx=5)
        lbl_pm_item_notes_value = ttk.Label(customer_notes_frame, text=LIMSVarConfig.external_customer_pm_item_notes,
                                            font=('arial', 12))
        lbl_pm_item_notes_value.grid(row=7, column=1)

        lbl_pm_item_notes_value_1 = ttk.Label(customer_notes_frame,
                                              text=LIMSVarConfig.external_customer_pm_item_notes_1, font=('arial', 12))
        lbl_pm_item_notes_value_1.grid(row=7, column=2)

        lbl_pm_item_notes_value_2 = ttk.Label(customer_notes_frame,
                                              text=LIMSVarConfig.external_customer_pm_item_notes_2, font=('arial', 12))
        lbl_pm_item_notes_value_2.grid(row=7, column=3)

        lbl_pm_item_notes_value_3 = ttk.Label(customer_notes_frame,
                                              text=LIMSVarConfig.external_customer_pm_item_notes_3, font=('arial', 12))
        lbl_pm_item_notes_value_3.grid(pady=3, row=8, column=1)

        lbl_pm_item_notes_value_4 = ttk.Label(customer_notes_frame,
                                              text=LIMSVarConfig.external_customer_pm_item_notes_4, font=('arial', 12))
        lbl_pm_item_notes_value_4.grid(row=8, column=2)

        lbl_pm_item_notes_value_5 = ttk.Label(customer_notes_frame,
                                              text=LIMSVarConfig.external_customer_pm_item_notes_5, font=('arial', 12))
        lbl_pm_item_notes_value_5.grid(row=8, column=3)

        # .............................Buttons....................................#

        btn_import_customer_information = ttk.Button(customer_information_frame, text="Import Customer Information",
                                                     width=30,
                                                     command=lambda: self.external_customer_information_data_load())
        btn_import_customer_information.bind("<Return>", lambda event: self.external_customer_information_data_load())
        btn_import_customer_information.grid(pady=5, row=1, column=2, columnspan=2)

        btn_backout_external_customer = ttk.Button(CalCert, text="Back", width=20,
                                                   command=lambda: self.sales_order_import_option_window(CalCert))
        btn_backout_external_customer.bind("<Return>", lambda event: self.sales_order_import_option_window(CalCert))
        btn_backout_external_customer.grid(pady=5, row=19, column=1, columnspan=2)

        btn_verify_external_customer = ttk.Button(CalCert, text="Next", width=20,
                                                  command=lambda: self.verify_external_customer_information_input())
        btn_verify_external_customer.bind("<Return>", lambda event: self.verify_external_customer_information_input())
        btn_verify_external_customer.grid(pady=5, row=19, column=3, columnspan=2)

    # -----------------------------------------------------------------------#

    # Command to ensure that fields have been filled prior to advancing to next window
    def verify_external_customer_information_input(self):
        if external_customer_sales_order_number_value.get() == "":
            tm.showerror("Error", 'Please complete the required fields!')
        else:
            self.device_instrument_description_information(CalCert)

    # -----------------------------------------------------------------------#

    # Import Command from AS400 Text Files
    def external_customer_information_data_load(self):

        LIMSVarConfig.external_customer_sales_order_number_helper = external_customer_sales_order_number_value.get()

        from LIMSAS400Search import AppAS400ExecutableHelper
        perform_customer_information_import = AppAS400ExecutableHelper()
        perform_customer_information_import.customer_information_helper()

        # Customer Name
        lbl_customer_name_value.config(text=LIMSVarConfig.external_customer_name)

        # Customer Address
        lbl_customer_address_value.config(text=LIMSVarConfig.external_customer_address)
        lbl_customer_address_value_1.config(text=LIMSVarConfig.external_customer_address_1)
        lbl_customer_address_value_2.config(text=LIMSVarConfig.external_customer_address_2)

        # Customer City
        lbl_customer_city_value.config(text=LIMSVarConfig.external_customer_city)

        # Customer State,Country and Zip Code
        lbl_state_country_zip_value.config(text=LIMSVarConfig.external_customer_state_country_zip)

        # Customer PO
        lbl_customer_purchase_order_value.config(text=LIMSVarConfig.external_customer_po)

        # Customer RMA
        lbl_customer_rma_value.config(text=LIMSVarConfig.external_customer_rma_number)

        # Customer Header Notes
        lbl_header_notes.config(text=LIMSVarConfig.external_customer_header_notes)
        lbl_header_notes_1.config(text=LIMSVarConfig.external_customer_header_notes_1)
        lbl_header_notes_2.config(text=LIMSVarConfig.external_customer_header_notes_2)
        lbl_header_notes_3.config(text=LIMSVarConfig.external_customer_header_notes_3)
        lbl_header_notes_4.config(text=LIMSVarConfig.external_customer_header_notes_4)
        lbl_header_notes_5.config(text=LIMSVarConfig.external_customer_header_notes_5)

        # Customer Item Notes
        lbl_item_notes.config(text=LIMSVarConfig.external_customer_item_notes)
        lbl_item_notes_1.config(text=LIMSVarConfig.external_customer_item_notes_1)
        lbl_item_notes_2.config(text=LIMSVarConfig.external_customer_item_notes_2)
        lbl_item_notes_3.config(text=LIMSVarConfig.external_customer_item_notes_3)
        lbl_item_notes_4.config(text=LIMSVarConfig.external_customer_item_notes_4)
        lbl_item_notes_5.config(text=LIMSVarConfig.external_customer_item_notes_5)

        # Customer Footer Notes
        lbl_footer_notes.config(text=LIMSVarConfig.external_customer_footer_notes)
        lbl_footer_notes_1.config(text=LIMSVarConfig.external_customer_footer_notes_1)
        lbl_footer_notes_2.config(text=LIMSVarConfig.external_customer_footer_notes_2)
        lbl_footer_notes_3.config(text=LIMSVarConfig.external_customer_footer_notes_3)
        lbl_footer_notes_4.config(text=LIMSVarConfig.external_customer_footer_notes_4)
        lbl_footer_notes_5.config(text=LIMSVarConfig.external_customer_footer_notes_5)

        # Customer PM Item Notes
        lbl_pm_item_notes_value.config(text=LIMSVarConfig.external_customer_pm_item_notes)
        lbl_pm_item_notes_value_1.config(text=LIMSVarConfig.external_customer_pm_item_notes_1)
        lbl_pm_item_notes_value_2.config(text=LIMSVarConfig.external_customer_pm_item_notes_2)
        lbl_pm_item_notes_value_3.config(text=LIMSVarConfig.external_customer_pm_item_notes_3)
        lbl_pm_item_notes_value_4.config(text=LIMSVarConfig.external_customer_pm_item_notes_4)
        lbl_pm_item_notes_value_5.config(text=LIMSVarConfig.external_customer_pm_item_notes_5)

    # -----------------------------------------------------------------------#

    # Command to Send User to Equipment Description Section of Cal Certificate
    def device_instrument_description_information(self, window):
        window.withdraw()

        global CertDUTDetails, certificate_type, dut_date_received, dut_date_of_calibration, dut_calibration_due_date, \
            dut_model_number, device_certificate_of_calibration_number, dut_condition, dut_output_type, \
            dut_instrument_identification_number, device_customer_identification_number, dut_date_code, \
            dut_model_number, btn_import_certificate_number, btn_import_cert_serial_number, btn_device_specification, \
            device_information_frame, device_instrument_identification

        device_received_date = StringVar()
        device_date_of_calibration = StringVar()
        device_calibration_due_date = StringVar()
        device_model_number = StringVar()
        device_instrument_identification = StringVar()
        device_customer_identification = StringVar()
        device_date_code = StringVar()

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        from LIMSCCCreation import AppCCModule
        acc = AppCommonCommands()
        ahw = AppHelpWindows()
        accm = AppCCModule()

        # ........................Main Window Properties.........................#

        CertDUTDetails = Toplevel()
        CertDUTDetails.title("Certificate of Calibration Creation")
        CertDUTDetails.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 415
        height = 505
        screen_width = CertDUTDetails.winfo_screenwidth()
        screen_height = CertDUTDetails.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        CertDUTDetails.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CertDUTDetails.focus_force()
        CertDUTDetails.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CertDUTDetails))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(CertDUTDetails, width, height, x, y)

        if LIMSVarConfig.customer_selection_check == int(0):
            menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CertDUTDetails)),
                                               ("Customer Information",
                                                lambda: self.external_customer_information(CertDUTDetails)),
                                               ("Logout", lambda: acc.software_signout(CertDUTDetails)),
                                               ("Quit", lambda: acc.software_close(CertDUTDetails))])
            menubar.add_menu("Help", commands=[("Help", lambda: ahw.device_instrument_information_description_help())])
        else:
            menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CertDUTDetails)),
                                               ("Internal Customer Information",
                                                lambda: accm.internal_customer_selection(CertDUTDetails)),
                                               ("Logout", lambda: acc.software_signout(CertDUTDetails)),
                                               ("Quit", lambda: acc.software_close(CertDUTDetails))])
            menubar.add_menu("Help", commands=[("Help", lambda: ahw.device_instrument_information_description_help())])

        # ...........................Frame Creation..............................#

        device_date_frame = LabelFrame(CertDUTDetails, text="Calibration Certificate Information", relief=SOLID, bd=1,
                                       labelanchor="n")
        device_date_frame.grid(row=0, column=0, rowspan=4, columnspan=2, padx=8, pady=5)

        device_information_frame = LabelFrame(CertDUTDetails, text="Instrument Information", relief=SOLID, bd=1,
                                              labelanchor="n")
        device_information_frame.grid(row=5, column=0, rowspan=7, columnspan=2, padx=8, pady=5)

        # .........................Labels and Entries............................#

        # Ask for Certificate Type / Option
        lbl_certificate_type = ttk.Label(device_date_frame, text="Certificate Type:", font=('arial', 12),
                                         anchor="n")
        lbl_certificate_type.grid(row=0, pady=5)
        certificate_type = ttk.Combobox(device_date_frame, values=[" ", "NIST", "17025 w/ EMU",
                                                                   "17025 w/ EMU & TUR"])
        certificate_type.config(state="active", width=len("17025 w/ EMU & TUR")+2)
        certificate_type.grid(row=0, column=2)
        acc.always_active_style(certificate_type)
        certificate_type.focus()

        # Ask for Date DUT was Received
        lbl_dut_date_received = ttk.Label(device_date_frame, text="Date Received:", font=('arial', 12), anchor="n")
        lbl_dut_date_received.grid(row=1, pady=5, padx=5)
        dut_date_received = ttk.Entry(device_date_frame, textvariable=device_received_date, font=('arial', 12))
        dut_date_received.grid(row=1, column=2)
        dut_date_received.config(width=15)
        dut_date_received.focus()

        # Ask for Date DUT was calibrated
        lbl_dut_date_of_calibration = ttk.Label(device_date_frame, text="Date of Calibration:", font=('arial', 12),
                                                anchor="n")
        lbl_dut_date_of_calibration.grid(row=2, pady=5, padx=5)
        dut_date_of_calibration = ttk.Entry(device_date_frame, textvariable=device_date_of_calibration,
                                            font=('arial', 12))
        dut_date_of_calibration.grid(row=2, column=2)
        dut_date_of_calibration.config(width=15)

        # Ask for Date DUT will be calibrated again
        lbl_dut_calibration_due_date = ttk.Label(device_date_frame, text="Calibration Due Date:", font=('arial', 12),
                                                 anchor="n")
        lbl_dut_calibration_due_date.grid(row=3, pady=5, padx=5)
        dut_calibration_due_date = ttk.Entry(device_date_frame, textvariable=device_calibration_due_date,
                                             font=('arial', 12))
        dut_calibration_due_date.grid(row=3, column=2)
        dut_calibration_due_date.config(width=15)

        # DUT certificate of calibration number displayed upon import
        lbl_certificate_of_calibration_number = ttk.Label(device_date_frame, text="Certificate Number:",
                                                          font=('arial', 12), anchor="n")
        lbl_certificate_of_calibration_number.grid(row=4, pady=5, padx=5)
        device_certificate_of_calibration_number = ttk.Label(device_date_frame,
                                                             text=LIMSVarConfig.certificate_of_calibration_number,
                                                             font=('arial', 12),
                                                             anchor="n")
        device_certificate_of_calibration_number.grid(row=4, column=2)
        device_certificate_of_calibration_number.config(width=15)

        btn_import_certificate_number = ttk.Button(device_date_frame, text="Import Certificate #", width=20,
                                                   command=lambda: self.import_device_certificate_of_calibration_number())
        btn_import_certificate_number.bind("<Return>",
                                           lambda event: self.import_device_certificate_of_calibration_number())
        btn_import_certificate_number.grid(pady=5, row=5, column=0, columnspan=1)

        btn_import_cert_serial_number = ttk.Button(device_date_frame, text="Import Cert. & Serial #", width=20,
                                                   command=lambda: self.import_device_certificate_and_serial_number())
        btn_import_cert_serial_number.bind("<Return>",
                                           lambda event: self.import_device_certificate_and_serial_number())
        btn_import_cert_serial_number.grid(pady=5, row=5, column=2, columnspan=1)

        # Ask for Condition of DUT being Calibrated
        lbl_dut_condition = ttk.Label(device_information_frame, text="Condition of DUT:", font=('arial', 12),
                                      anchor="n")
        lbl_dut_condition.grid(row=1, pady=5)
        dut_condition = ttk.Combobox(device_information_frame, values=[" ", "New", "Used", "Repaired"])
        dut_condition.config(state="active", width=18)
        dut_condition.grid(row=1, column=2)
        acc.always_active_style(dut_condition)

        # Ask for Output Type of Instrument being Calibrated
        lbl_output_type = ttk.Label(device_information_frame, text="Output Type:", font=('arial', 12))
        lbl_output_type.grid(row=2, pady=5)
        dut_output_type = ttk.Combobox(device_information_frame, values=[" ", "Single", "Dual", "Transmitter"])
        dut_output_type.config(state="active", width=18)
        dut_output_type.grid(row=2, column=2)
        acc.always_active_style(dut_output_type)

        # Ask for Model Number of Instrument being Calibrated
        lbl_dut_model_number = ttk.Label(device_information_frame, text="Model Number:", font=('arial', 12))
        lbl_dut_model_number.grid(row=3, pady=5)
        dut_model_number = ttk.Entry(device_information_frame, textvariable=device_model_number, font=('arial', 12))
        dut_model_number.grid(row=3, column=2)
        dut_model_number.config(width=15)

        # Ask for Date Code of Instrument being Calibrated
        lbl_dut_date_code = ttk.Label(device_information_frame, text="Date Code:", font=('arial', 12))
        lbl_dut_date_code.grid(row=4, pady=5)
        dut_date_code = ttk.Entry(device_information_frame, textvariable=device_date_code, font=('arial', 12))
        dut_date_code.grid(row=4, column=2)
        dut_date_code.config(width=15)

        # Ask for Customer Asset ID Number of Instrument being Calibrated
        lbl_dut_customer_identification = ttk.Label(device_information_frame, text="Customer Instrument ID Number:",
                                                    font=('arial', 12))
        lbl_dut_customer_identification.grid(row=5, pady=5, padx=5)
        device_customer_identification_number = ttk.Entry(device_information_frame,
                                                          textvariable=device_customer_identification,
                                                          font=('arial', 12))
        device_customer_identification_number.grid(row=5, column=2)
        device_customer_identification_number.config(width=15)

        # Ask for Asset ID Number of Instrument being Calibrated
        lbl_dut_identification = ttk.Label(device_information_frame, text="Instrument ID Number:", font=('arial', 12))
        lbl_dut_identification.grid(row=6, pady=5)
        dut_instrument_identification_number = ttk.Entry(device_information_frame,
                                                         textvariable=device_instrument_identification,
                                                         font=('arial', 12))
        dut_instrument_identification_number.grid(row=6, column=2)
        dut_instrument_identification_number.config(width=15)

        # This label is a dummy label that does nothing but help format the window
        dummy = ttk.Label(device_date_frame)
        dummy.grid(row=1, column=3)
        dummy.config(width=2)
        dummy1 = ttk.Label(device_information_frame)
        dummy1.grid(row=1, column=3)
        dummy1.config(width=2)

        # ..............................Buttons...................................#

        if LIMSVarConfig.customer_selection_check == int(0):
            btn_back_out_instrument_description = ttk.Button(CertDUTDetails, text="Back", width=20,
                                                             command=lambda: self.external_customer_information(CertDUTDetails))
            btn_back_out_instrument_description.bind("<Return>",
                                                     lambda event: self.external_customer_information(CertDUTDetails))
            btn_back_out_instrument_description.grid(pady=5, row=17, column=0, columnspan=1)
        else:
            btn_back_out_instrument_description = ttk.Button(CertDUTDetails, text="Back", width=20,
                                                             command=lambda: accm.internal_customer_selection(CertDUTDetails))
            btn_back_out_instrument_description.bind("<Return>",
                                                     lambda event: accm.internal_customer_selection(CertDUTDetails))
            btn_back_out_instrument_description.grid(pady=5, row=17, column=0, columnspan=1)

        btn_device_specification = ttk.Button(CertDUTDetails, text="Next", width=20,
                                              command=lambda: self.device_certificate_and_information_input_verify())
        btn_device_specification.bind("<Return>", lambda event: self.device_certificate_and_information_input_verify())
        btn_device_specification.grid(pady=5, row=17, column=1, columnspan=1)

    # -----------------------------------------------------------------------#

    # Command to Search for Next Available Cert Number
    def import_device_certificate_of_calibration_number(self):
        global dut_instrument_identification_number

        btn_import_certificate_number.config(cursor="watch")
        LIMSVarConfig.certificate_of_calibration_number = ""
        LIMSVarConfig.device_serial_number = ""
        device_certificate_of_calibration_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
        dut_instrument_identification_number.destroy()

        if LIMSVarConfig.customer_selection_check == int(0):
            LIMSVarConfig.external_customer_sales_order_number_helper = external_customer_sales_order_number_value.get()
        else:
            LIMSVarConfig.external_customer_name = "Dwyer Instruments, Inc."
            LIMSVarConfig.external_customer_sales_order_number_helper = "-"
            LIMSVarConfig.external_customer_rma_number = "-"

        LIMSVarConfig.calibration_date_helper = dut_date_of_calibration.get()
        LIMSVarConfig.calibration_due_date_helper = dut_calibration_due_date.get()

        try:
            year = datetime.today().year
            excel_database = open("\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls" %(year, year))
            if excel_database.closed is False:
                excel_database.close()
                dut_instrument_identification_number = ttk.Entry(device_information_frame,
                                                                 textvariable=device_instrument_identification,
                                                                 font=('arial', 12))
                dut_instrument_identification_number.grid(row=6, column=2)
                dut_instrument_identification_number.config(width=15)
                from LIMSCertDBase import AppCertificateDatabase
                apd = AppCertificateDatabase()
                print('Running certificate_number_checker')
                apd.certificate_number_checker()
                if LIMSVarConfig.certificate_of_calibration_number != "":
                    device_certificate_of_calibration_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                else:
                    apd.certificate_number_helper()
                    device_certificate_of_calibration_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                btn_import_certificate_number.config(cursor="arrow")
        except IOError as e:
            btn_import_certificate_number.config(cursor="arrow")
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------#

    # Command to Search for Next Available Cert and Serial Number
    def import_device_certificate_and_serial_number(self):
        global dut_instrument_identification_number

        btn_import_cert_serial_number.config(cursor="watch")
        LIMSVarConfig.certificate_of_calibration_number = ""
        LIMSVarConfig.device_serial_number = ""
        device_certificate_of_calibration_number.config(text=LIMSVarConfig.certificate_of_calibration_number)

        if LIMSVarConfig.customer_selection_check == int(0):
            LIMSVarConfig.external_customer_sales_order_number_helper = external_customer_sales_order_number_value.get()
        else:
            LIMSVarConfig.external_customer_name = "Dwyer Instruments, Inc."
            LIMSVarConfig.external_customer_sales_order_number_helper = "-"
            LIMSVarConfig.external_customer_rma_number = "-"

        LIMSVarConfig.calibration_date_helper = dut_date_of_calibration.get()
        LIMSVarConfig.calibration_due_date_helper = dut_calibration_due_date.get()

        try:
            print('getting year and opening cert database')
            year = datetime.today().year
            excel_database = open("\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls" % (year, year))
            print('Database opened')
            if excel_database.closed is False:
                excel_database.close()
                dut_instrument_identification_number.destroy()
                from LIMSCertDBase import AppCertificateDatabase
                apd = AppCertificateDatabase()
                apd.certificate_serial_number_checker()
                if LIMSVarConfig.certificate_of_calibration_number != "" and LIMSVarConfig.device_serial_number != "":
                    device_certificate_of_calibration_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                    dut_instrument_identification_number = ttk.Label(device_information_frame,
                                                                     text=LIMSVarConfig.device_serial_number,
                                                                     font=('arial', 12))
                    dut_instrument_identification_number.grid(row=6, column=2)
                    dut_instrument_identification_number.config(width=15)
                else:
                    apd.certificate_serial_number_helper()
                    device_certificate_of_calibration_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                    dut_instrument_identification_number = ttk.Label(device_information_frame,
                                                                     text=LIMSVarConfig.device_serial_number,
                                                                     font=('arial', 12))
                    dut_instrument_identification_number.grid(row=6, column=2)
                    dut_instrument_identification_number.config(width=15)
                btn_import_cert_serial_number.config(cursor="arrow")
        except IOError as e:
            btn_import_cert_serial_number.config(cursor="arrow")
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------#

    # Command to ensure that fields have been filled prior to advancing to next window
    def device_certificate_and_information_input_verify(self):

        if LIMSVarConfig.device_serial_number == "":
            if certificate_type.get() == "" or dut_date_received.get() == "" or dut_date_of_calibration.get() == "" or \
                    dut_calibration_due_date.get() == "" or LIMSVarConfig.certificate_of_calibration_number == "" or \
                    dut_condition.get() == " " or dut_output_type.get() == " " or \
                    dut_instrument_identification_number.get() == "" or \
                    device_customer_identification_number.get() == "" or dut_date_code.get() == "" or \
                    dut_model_number.get() == "":
                tm.showerror("Error", 'Please complete the required fields!')
            else:
                self.device_under_test_calibration_specifications(CertDUTDetails)
        elif LIMSVarConfig.device_serial_number != "":
            if certificate_type.get() == "" or dut_date_received.get() == "" or dut_date_of_calibration.get() == "" or \
                    dut_calibration_due_date.get() == "" or LIMSVarConfig.certificate_of_calibration_number == "" or \
                    dut_condition.get() == " " or dut_output_type.get() == " " or \
                    device_customer_identification_number.get() == "" or dut_date_code.get() == "" or \
                    dut_model_number.get() == "":
                tm.showerror("Error", 'Please complete the required fields!')
            else:
                self.device_under_test_calibration_specifications(CertDUTDetails)

    # -----------------------------------------------------------------------#

    # Command to Send User to Calibration Details Section of Cal Certificate
    def device_under_test_calibration_specifications(self, window):

        # btn_device_specification.config(cursor="watch")
        LIMSVarConfig.certificate_option = certificate_type.get()
        LIMSVarConfig.device_date_received_helper = dut_date_received.get()
        LIMSVarConfig.device_model_number_helper = dut_model_number.get()
        if LIMSVarConfig.device_serial_number == "":
            LIMSVarConfig.device_identification_number_helper = dut_instrument_identification_number.get()
        elif LIMSVarConfig.device_serial_number != "":
            LIMSVarConfig.device_identification_number_helper = LIMSVarConfig.device_serial_number
        LIMSVarConfig.device_date_code_helper = dut_date_code.get()

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        aph = AppHelpWindows()

        global DUTCalDetails, dut_output_type, tab1, tab2, tab3, error_type_width, \
            discipline_and_unit_frame, measurement_discipline_list, measurement_discipline_list_1, \
            measurement_discipline_type_width, measurement_discipline_type, measurement_discipline_type_1, \
            measurement_discipline_selection, measurement_discipline_selection_1, units_of_measure_list, \
            units_of_measure_list_1, device_under_test_measurement_units_value, \
            device_under_test_measurement_units_value_1, device_under_test_units, device_under_test_units_1, \
            device_under_test_minimum_full_scale, device_under_test_minimum_full_scale_1, \
            device_under_test_maximum_full_scale, device_under_test_maximum_full_scale_1, \
            reference_resolution_number, reference_resolution_number_1, reference_resolution_value, \
            reference_resolution_value_1, device_under_test_resolution, device_under_test_resolution_1, \
            device_under_test_number_of_test_points, device_under_test_number_of_test_points_1, \
            number_of_test_points, number_of_test_points_1, device_under_test_test_point_direction, \
            device_under_test_test_point_direction_1, test_point_direction, test_point_direction_1, \
            first_operator_list, supplemental_operator_list, selected_operator, selected_operator_1, \
            selected_operator_2, selected_operator_3, measurement_error_assignment_type, \
            measurement_error_assignment_type_1, measurement_error_assignment_type_2, \
            measurement_error_assignment_type_3, operator_drop_down, operator_drop_down_1, \
            operator_drop_down_2, operator_drop_down_3, measurement_error_type_1_value, \
            measurement_error_type_1_value_additional, measurement_error_type_1_value_additional_1, \
            measurement_error_type_1_value_additional_2, device_under_test_error_selection, \
            device_under_test_error_selection_1, device_under_test_error_selection_2, \
            device_under_test_error_selection_3, second_selected_operator, second_selected_operator_1, \
            second_selected_operator_2, second_selected_operator_3, second_measurement_error_assignment_type, \
            second_measurement_error_assignment_type_1, second_measurement_error_assignment_type_2, \
            second_measurement_error_assignment_type_3, second_operator_drop_down, \
            second_operator_drop_down_1, second_operator_drop_down_2, second_operator_drop_down_3, \
            second_measurement_error_type_value, second_measurement_error_type_value_1, \
            second_measurement_error_type_value_2, second_measurement_error_type_value_3, \
            second_device_under_test_error_selection, second_device_under_test_error_selection_1, \
            second_device_under_test_error_selection_2, second_device_under_test_error_selection_3, device_accuracy_frame_1, device_accuracy_frame, dut_test_step_direction_frame_1, dut_test_step_direction_frame, device_measurement_resolution_frame_1, device_measurement_resolution_frame, device_span_frame_1, device_span_frame, error_type_drop_down_list

        device_under_test_resolution_value = StringVar()
        device_under_test_resolution_value_1 = StringVar()
        device_under_test_minimum_full_scale_value = StringVar()
        device_under_test_minimum_full_scale_value_1 = StringVar()
        device_under_test_maximum_full_scale_value = StringVar()
        device_under_test_maximum_full_scale_value_1 = StringVar()
        device_under_test_measurement_units_value = StringVar()
        device_under_test_measurement_units_value_1 = StringVar()
        device_under_test_error_selection_value = StringVar()
        device_under_test_error_selection_value_1 = StringVar()
        device_under_test_error_selection_value_2 = StringVar()
        device_under_test_error_selection_value_3 = StringVar()
        device_under_test_second_error_selection_value = StringVar()
        device_under_test_second_error_selection_value_1 = StringVar()
        device_under_test_second_error_selection_value_2 = StringVar()
        device_under_test_second_error_selection_value_3 = StringVar()

        # ........................Main Window Properties..........................#

        if len(LIMSVarConfig.calibration_equipment_reference_array) is None or \
                len(LIMSVarConfig.calibration_equipment_reference_array) == 0:
            from LIMSRefStd import AppReferenceStandardDatabase
            arsd = AppReferenceStandardDatabase()
            arsd.calibration_equipment_database_csv()
        else:
            pass

        btn_device_specification.config(cursor="arrow")

        window.withdraw()
        DUTCalDetails = Toplevel()
        DUTCalDetails.title("DUT Calibration Parameters")
        DUTCalDetails.iconbitmap("Required Images\\DwyerLogo.ico")

        if dut_output_type.get() == "Single":
            height = 570
            width = 655
        elif dut_output_type.get() == "Transmitter":
            height = 690
            width = 655
        else:
            height = 600
            width = 1250

        screen_width = DUTCalDetails.winfo_screenwidth()
        screen_height = DUTCalDetails.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        DUTCalDetails.geometry("%dx%d+%d+%d" % (width, height, x, y))
        DUTCalDetails.focus_force()
        DUTCalDetails.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(DUTCalDetails))

        # ..........................Window Header.................................#

        lbl_calibration_parameters_frame = Label(DUTCalDetails, text="DUT Calibration Parameters", width=20,
                                                 font=('arial', 12), bd=10)
        lbl_calibration_parameters_frame.grid(row=1, column=3, columnspan=1)

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(DUTCalDetails, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(DUTCalDetails)),
                                           ("DUT Information",
                                            lambda: self.device_instrument_description_information(DUTCalDetails)),
                                           ("Logout", lambda: acc.software_signout(DUTCalDetails)),
                                           ("Quit", lambda: acc.software_close(DUTCalDetails))])
        menubar.add_menu("Help", commands=[("Help", lambda: aph.device_under_test_specification_details())])

        # .........................Notebook Creation..............................#

        notebook_frame = ttk.Notebook(DUTCalDetails)
        notebook_frame.grid(row=2, column=0, sticky=E + W + N + S, padx=8, pady=5, columnspan=7)

        tab1 = Frame(notebook_frame)
        tab2 = Frame(notebook_frame)
        tab3 = Frame(notebook_frame)

        if dut_output_type.get() == "Single":
            notebook_frame.add(tab1, text="Single", compound=TOP)
            notebook_frame.add(tab2, text="Dual", compound=TOP, state="disabled")
            notebook_frame.add(tab3, text="Transmitter", compound=TOP, state="disabled")
        elif dut_output_type.get() == "Dual":
            notebook_frame.add(tab1, text="Single", compound=TOP, state="disabled")
            notebook_frame.add(tab2, text="Dual", compound=TOP)
            notebook_frame.add(tab3, text="Transmitter", compound=TOP, state="disabled")
        elif dut_output_type.get() == "Transmitter":
            notebook_frame.add(tab1, text="Single", compound=TOP, state="disabled")
            notebook_frame.add(tab2, text="Dual", compound=TOP, state="disabled")
            notebook_frame.add(tab3, text="Transmitter", compound=TOP)

        # ..................Error Type Label & Drop down...........................#

        if dut_output_type.get() != "Transmitter":
            error_type_drop_down_list = [" ", "Actual Error", '% of Full Scale', '% of FS (Range)',
                                         '% of Reading']
        elif dut_output_type.get() == "Transmitter":
            error_type_drop_down_list = [" ", "Actual Error", '% of Full Scale', '% of Reading']
        error_type_width = len(max(error_type_drop_down_list, key=len))

        # ///////////////////////////////////////////////////////////////////////#
        # The following is information provided on each of the output type tabs
        # in the calibration parameters window. Though repetitive, each has
        # its own functionality when generating the following table like
        # data acquisition phase.
        # ///////////////////////////////////////////////////////////////////////#

        # ..............Single Output Discipline Label and Dropdown...............#

        # Frame Creation for Full Scale Range of DUT
        if dut_output_type.get() == "Single":
            discipline_and_unit_frame = LabelFrame(tab1, text="Discipline and Unit Selection", relief=SOLID,
                                                   bd=1, labelanchor="n")
            discipline_and_unit_frame.grid(pady=5, row=0, column=0, rowspan=2, columnspan=8)
        if dut_output_type.get() == "Dual":
            discipline_and_unit_frame = LabelFrame(tab2, text="Discipline and Unit Selection", relief=SOLID,
                                                   bd=1, labelanchor="n")
            discipline_and_unit_frame.grid(pady=5, row=0, column=0, rowspan=2, columnspan=9)
        if dut_output_type.get() == "Transmitter":
            discipline_and_unit_frame = LabelFrame(tab3, text="Discipline, Transmitter Type and Unit Selection",
                                                   relief=SOLID, bd=1, labelanchor="n")
            discipline_and_unit_frame.grid(pady=5, row=0, column=0, rowspan=2, columnspan=8)

        # Measurement Discipline Designation
        measurement_discipline_list = [" ", "Capacitance", "Current", "Concentration", "Flow", "Frequency", "Humidity",
                                       "Pressure", "Resistance", "Temperature", "Velocity", "Voltage"]
        measurement_discipline_type_width = len(max(measurement_discipline_list, key=len))
        measurement_discipline_type = StringVar(DUTCalDetails)

        if dut_output_type.get() == "Single":
            lbl_discipline = ttk.Label(discipline_and_unit_frame, text="Discipline:", font=('arial', 12))
            lbl_discipline.grid(row=0, column=1)
        else:
            lbl_discipline = ttk.Label(discipline_and_unit_frame, text="Discipline 1:", font=('arial', 12))
            lbl_discipline.grid(row=0, column=1)
        measurement_discipline_selection = ttk.Combobox(discipline_and_unit_frame,
                                                        values=measurement_discipline_list)
        acc.always_active_style(measurement_discipline_selection)
        measurement_discipline_selection.config(state="active", width=error_type_width)
        measurement_discipline_selection.focus()
        measurement_discipline_selection.grid(pady=5, row=0, column=2, columnspan=1)

        # Update unit list based on measurement discipline selected
        if dut_output_type.get() == "Single":
            btn_update_device_calibration_specifications = ttk.Button(discipline_and_unit_frame,
                                                                      text="Update Units", width=20,
                                                                      command=lambda: self.update_dut_measurement_unit_selection())
            btn_update_device_calibration_specifications.bind("<Return>",
                                                              lambda event: self.update_dut_measurement_unit_selection())
            btn_update_device_calibration_specifications.grid(pady=5, padx=5, row=0, column=3, columnspan=1)
        else:
            btn_update_device_calibration_specifications = ttk.Button(discipline_and_unit_frame,
                                                                      text="Update Units", width=20,
                                                                      command=lambda: self.update_dut_measurement_unit_selection())
            btn_update_device_calibration_specifications.bind("<Return>",
                                                              lambda event: self.update_dut_measurement_unit_selection())
            btn_update_device_calibration_specifications.grid(pady=5, padx=5, row=0, column=3, columnspan=1,
                                                              rowspan=2)

        # Units of Measure
        units_of_measure_list = [" "]
        device_under_test_measurement_units_value = StringVar(DUTCalDetails)
        lbl_dut_measurement_units = ttk.Label(discipline_and_unit_frame, text="Units:", font=('arial', 12))
        lbl_dut_measurement_units.grid(row=0, column=4)
        device_under_test_units = ttk.Combobox(discipline_and_unit_frame, values=units_of_measure_list)
        acc.always_active_style(device_under_test_units)
        device_under_test_units.config(state="active", width=error_type_width)
        device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        if dut_output_type.get() == "Dual":

            # Measurement Discipline Designation
            measurement_discipline_list_1 = [" ", "Capacitance", "Current", "Concentration", "Flow", "Frequency",
                                             "Humidity", "Pressure", "Resistance", "Temperature", "Velocity", "Voltage"]
            measurement_discipline_type_1 = StringVar(DUTCalDetails)
            lbl_discipline_1 = ttk.Label(discipline_and_unit_frame, text="Discipline 2:", font=('arial', 12))
            lbl_discipline_1.grid(row=1, column=1)
            measurement_discipline_selection_1 = ttk.Combobox(discipline_and_unit_frame,
                                                              values=measurement_discipline_list_1)
            acc.always_active_style(measurement_discipline_selection_1)
            measurement_discipline_selection_1.config(state="active", width=error_type_width)
            measurement_discipline_selection_1.grid(pady=5, padx=5, row=1, column=2, columnspan=1)

            # Units of Measure
            units_of_measure_list_1 = [" "]
            device_under_test_measurement_units_value_1 = StringVar(DUTCalDetails)
            lbl_dut_measurement_units_1 = ttk.Label(discipline_and_unit_frame, text="Units:", font=('arial', 12))
            lbl_dut_measurement_units_1.grid(row=1, column=4)
            device_under_test_units_1 = ttk.Combobox(discipline_and_unit_frame,
                                                     values=units_of_measure_list_1)
            acc.always_active_style(device_under_test_units_1)
            device_under_test_units_1.config(state="active", width=error_type_width)
            device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

        elif dut_output_type.get() == "Transmitter":

            # Transmitter Discipline Designation
            measurement_discipline_list_1 = [" ", "Current", "Voltage"]
            measurement_discipline_type_1 = StringVar(DUTCalDetails)
            lbl_discipline_1 = ttk.Label(discipline_and_unit_frame, text="Transmitter Type:",
                                         font=('arial', 12))
            lbl_discipline_1.grid(row=1, column=1)
            measurement_discipline_selection_1 = ttk.Combobox(discipline_and_unit_frame,
                                                              values=measurement_discipline_list_1)
            acc.always_active_style(measurement_discipline_selection_1)
            measurement_discipline_selection_1.config(state="active", width=error_type_width)
            measurement_discipline_selection_1.grid(pady=5, padx=5, row=1, column=2, columnspan=1)

            # Units of Measure
            units_of_measure_list_1 = [" "]
            device_under_test_measurement_units_value_1 = StringVar(DUTCalDetails)
            lbl_dut_measurement_units_1 = ttk.Label(discipline_and_unit_frame, text="Units:",
                                                    font=('arial', 12))
            lbl_dut_measurement_units_1.grid(row=1, column=4)
            device_under_test_units_1 = ttk.Combobox(discipline_and_unit_frame,
                                                     values=units_of_measure_list_1)
            acc.always_active_style(device_under_test_units_1)
            device_under_test_units_1.config(state="active", width=error_type_width)
            device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

        # These labels are dummy labels that do nothing but help format the frame
        dummy = ttk.Label(discipline_and_unit_frame)
        dummy.grid(row=0, column=0)
        dummy.config(width=1)
        dummy_1 = ttk.Label(discipline_and_unit_frame)
        dummy_1.grid(row=0, column=6)
        dummy.config(width=1)
        # ........................DUT Full Scale Range(s)..........................#

        # Frame Creation for Full Scale Range of DUT
        if dut_output_type.get() == "Single":
            device_span_frame = LabelFrame(tab1, text="DUT Full Scale Range", relief=SOLID, bd=1,
                                           labelanchor="n")
            device_span_frame.grid(pady=5, row=2, column=0, rowspan=2, columnspan=8)
        if dut_output_type.get() == "Dual":
            device_span_frame = LabelFrame(tab2, text="Discipline 1: DUT Full Scale Range", relief=SOLID, bd=1,
                                           labelanchor="n")
            device_span_frame.grid(pady=5, row=2, column=1, rowspan=2, columnspan=3)
            device_span_frame_1 = LabelFrame(tab2, text="Discipline 2: DUT Full Scale Range", relief=SOLID,
                                             bd=1, labelanchor="n")
            device_span_frame_1.grid(pady=5, row=2, column=5, rowspan=2, columnspan=3)
        if dut_output_type.get() == "Transmitter":
            device_span_frame = LabelFrame(tab3, text="DUT Full Scale Range", relief=SOLID, bd=1,
                                           labelanchor="n")
            device_span_frame.grid(pady=5, row=2, column=0, rowspan=2, columnspan=8)
            device_span_frame_1 = LabelFrame(tab3, text="Transmitter Span", relief=SOLID, bd=1, labelanchor="n")
            device_span_frame_1.grid(pady=5, row=4, column=0, rowspan=2, columnspan=8)

        # DUT Full Scale Range
        lbl_minimum_full_scale = ttk.Label(device_span_frame, text="Minimum Value:", font=('arial', 12),
                                           anchor="n")
        lbl_minimum_full_scale.grid(row=0, column=1, pady=5)
        device_under_test_minimum_full_scale = ttk.Entry(device_span_frame,
                                                         textvariable=device_under_test_minimum_full_scale_value,
                                                         font=('arial', 12))
        device_under_test_minimum_full_scale.grid(row=0, column=2, padx=5)
        device_under_test_minimum_full_scale.config(width=13)
        lbl_maximum_full_scale = ttk.Label(device_span_frame, text="Maximum Value:", font=('arial', 12),
                                           anchor="n")
        lbl_maximum_full_scale.grid(row=1, column=1, pady=5)
        device_under_test_maximum_full_scale = ttk.Entry(device_span_frame,
                                                         textvariable=device_under_test_maximum_full_scale_value,
                                                         font=('arial', 12))
        device_under_test_maximum_full_scale.config(width=13)
        device_under_test_maximum_full_scale.grid(row=1, column=2, padx=5)

        # These labels are dummy labels that do nothing but help format the frame
        dummy = ttk.Label(device_span_frame)
        dummy.grid(row=1, column=3)
        dummy.config(width=2)
        dummy1 = ttk.Label(device_span_frame)
        dummy1.grid(row=1, column=0)
        dummy1.config(width=2)

        if dut_output_type.get() == "Transmitter" or dut_output_type.get() == "Dual":
            # DUT Transmitter Output
            lbl_minimum_full_scale_1 = ttk.Label(device_span_frame_1, text="Minimum Value:", font=('arial', 12),
                                                 anchor="n")
            lbl_minimum_full_scale_1.grid(row=0, column=1, pady=5)
            device_under_test_minimum_full_scale_1 = ttk.Entry(device_span_frame_1,
                                                               textvariable=device_under_test_minimum_full_scale_value_1,
                                                               font=('arial', 12))
            device_under_test_minimum_full_scale_1.grid(row=0, column=2, padx=5)
            device_under_test_minimum_full_scale_1.config(width=13)
            lbl_maximum_full_scale_1 = ttk.Label(device_span_frame_1, text="Maximum Value:", font=('arial', 12),
                                                 anchor="n")
            lbl_maximum_full_scale_1.grid(row=1, column=1, pady=5)
            device_under_test_maximum_full_scale_1 = ttk.Entry(device_span_frame_1,
                                                               textvariable=device_under_test_maximum_full_scale_value_1,
                                                               font=('arial', 12))
            device_under_test_maximum_full_scale_1.config(width=13)
            device_under_test_maximum_full_scale_1.grid(row=1, column=2, padx=5)

            # These labels are dummy labels that do nothing but help format the frame
            dummyf1 = ttk.Label(device_span_frame_1)
            dummyf1.grid(row=1, column=3)
            dummyf1.config(width=2)
            dummyf1_1 = ttk.Label(device_span_frame_1)
            dummyf1_1.grid(row=1, column=0)
            dummyf1_1.config(width=2)

        # ...................Calibration Measurement Resolution....................#

        # Frame Creation for Measurement Resolution
        if dut_output_type.get() == "Single":
            device_measurement_resolution_frame = LabelFrame(tab1, text="Measurement Resolution", relief=SOLID,
                                                             bd=1, labelanchor="n")
            device_measurement_resolution_frame.grid(pady=5, row=4, column=0, rowspan=2, columnspan=8)
        if dut_output_type.get() == "Dual":
            device_measurement_resolution_frame = LabelFrame(tab2, text="Discipline 1: Measurement Resolution",
                                                             relief=SOLID, bd=1, labelanchor="n")
            device_measurement_resolution_frame.grid(pady=5, row=4, column=1, rowspan=2, columnspan=3)
            device_measurement_resolution_frame_1 = LabelFrame(tab2,
                                                               text="Discipline 2: Measurement Resolution",
                                                               relief=SOLID, bd=1, labelanchor="n")
            device_measurement_resolution_frame_1.grid(pady=5, row=4, column=5, rowspan=2, columnspan=3)
        if dut_output_type.get() == "Transmitter":
            device_measurement_resolution_frame = LabelFrame(tab3, text="Measurement Resolution", relief=SOLID,
                                                             bd=1, labelanchor="n")
            device_measurement_resolution_frame.grid(pady=5, row=6, column=0, rowspan=2, columnspan=8)

        # Reference Resolution Option
        resolution_option_list = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
        reference_resolution_number = StringVar(DUTCalDetails)
        lbl_reference_resolution = ttk.Label(device_measurement_resolution_frame,
                                             text="Ref. Decimals Displayed:", width=20, font=('arial', 12),
                                             anchor="n")
        lbl_reference_resolution.grid(row=0, column=1)
        reference_resolution_value = ttk.Combobox(device_measurement_resolution_frame,
                                                  values=resolution_option_list)
        acc.always_active_style(reference_resolution_value)
        reference_resolution_value.config(state="active", width=error_type_width)
        reference_resolution_value.grid(pady=5, row=0, column=2, columnspan=1)

        # DUT Resolution Entry
        lbl_dut_resolution = ttk.Label(device_measurement_resolution_frame, text="DUT Resolution:", width=20,
                                       font=('arial', 12), anchor="n")
        lbl_dut_resolution.grid(row=1, column=1, pady=5)
        device_under_test_resolution = ttk.Entry(device_measurement_resolution_frame,
                                                 textvariable=device_under_test_resolution_value,
                                                 font=('arial', 12))
        acc.always_active_style(device_under_test_resolution)
        device_under_test_resolution.grid(row=1, column=2)
        device_under_test_resolution.config(width=13)

        # These labels are dummy labels that do nothing but help format the frame
        dummymf = ttk.Label(device_measurement_resolution_frame)
        dummymf.grid(row=1, column=3)
        dummymf.config(width=2)
        dummymf_1 = ttk.Label(device_measurement_resolution_frame)
        dummymf_1.grid(row=1, column=0)
        dummymf_1.config(width=2)

        if dut_output_type.get() == "Dual":
            # Reference Resolution Option
            resolution_option_list_1 = ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
            reference_resolution_number_1 = StringVar(DUTCalDetails)
            lbl_reference_resolution_1 = ttk.Label(device_measurement_resolution_frame_1,
                                                   text="Ref. Decimals Displayed:", width=20,
                                                   font=('arial', 12), anchor="n")
            lbl_reference_resolution_1.grid(row=0, column=1)
            reference_resolution_value_1 = ttk.Combobox(device_measurement_resolution_frame_1,
                                                        values=resolution_option_list_1)
            acc.always_active_style(reference_resolution_value_1)
            reference_resolution_value_1.config(state="active", width=error_type_width)
            reference_resolution_value_1.grid(pady=5, row=0, column=2, columnspan=1)

            # DUT Resolution Entry
            lbl_dut_resolution_1 = ttk.Label(device_measurement_resolution_frame_1, text="DUT Resolution:",
                                             width=20, font=('arial', 12), anchor="n")
            lbl_dut_resolution_1.grid(row=1, column=1, pady=5)
            device_under_test_resolution_1 = ttk.Entry(device_measurement_resolution_frame_1,
                                                       textvariable=device_under_test_resolution_value_1,
                                                       font=('arial', 12))
            acc.always_active_style(device_under_test_resolution_1)
            device_under_test_resolution_1.grid(row=1, column=2)
            device_under_test_resolution_1.config(width=13)

            # These labels are dummy labels that do nothing but help format the frame
            dummymf1 = ttk.Label(device_measurement_resolution_frame_1)
            dummymf1.grid(row=1, column=3)
            dummymf1.config(width=2)
            dummymf1_1 = ttk.Label(device_measurement_resolution_frame_1)
            dummymf1_1.grid(row=1, column=0)
            dummymf1_1.config(width=2)

        # .............DUT Number of Test Points and Type of Test Points.........#

        # Frame Creation for Test Steps and Direction
        if dut_output_type.get() == "Single":
            dut_test_step_direction_frame = LabelFrame(tab1, text="DUT Test Step & Direction", relief=SOLID,
                                                       bd=1, labelanchor="n")
            dut_test_step_direction_frame.grid(pady=5, row=6, column=0, rowspan=2, columnspan=8)
        elif dut_output_type.get() == "Dual":
            dut_test_step_direction_frame = LabelFrame(tab2, text="Discipline 1: DUT Test Step & Direction",
                                                       relief=SOLID, bd=1, labelanchor="n")
            dut_test_step_direction_frame.grid(pady=5, row=6, column=1, rowspan=2, columnspan=3)
            dut_test_step_direction_frame_1 = LabelFrame(tab2, text="Discipline 2: DUT Test Step & Direction",
                                                         relief=SOLID, bd=1, labelanchor="n")
            dut_test_step_direction_frame_1.grid(pady=5, row=6, column=5, rowspan=2, columnspan=3)
        elif dut_output_type.get() == "Transmitter":
            dut_test_step_direction_frame = LabelFrame(tab3, text="DUT Test Step & Direction", relief=SOLID,
                                                       bd=1, labelanchor="n")
            dut_test_step_direction_frame.grid(pady=5, row=8, column=0, rowspan=2, columnspan=8)

        # Test Step Number Designation
        if dut_output_type.get() == "Dual":
            test_point_number_list = [" ", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
        else:
            test_point_number_list = [" ", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "21"]
        device_under_test_number_of_test_points = StringVar(DUTCalDetails)
        lbl_number_of_test_points = ttk.Label(dut_test_step_direction_frame, text="Number of Test Points:",
                                              width=20, font=('arial', 12), anchor="n")
        lbl_number_of_test_points.grid(row=0, column=1)
        number_of_test_points = ttk.Combobox(dut_test_step_direction_frame,
                                             values=test_point_number_list)
        acc.always_active_style(number_of_test_points)
        number_of_test_points.config(state="active", width=error_type_width)
        number_of_test_points.grid(pady=5, row=0, column=2, columnspan=1)

        # Test Point Direction Assignment
        test_point_direction_list = [" ", "Ascending", "Descending", "Asc. & Desc."]
        device_under_test_test_point_direction = StringVar(DUTCalDetails)
        lbl_test_point_direction = ttk.Label(dut_test_step_direction_frame, text="Direction:", width=20,
                                             font=('arial', 12), anchor="n")
        lbl_test_point_direction.grid(row=1, column=1)
        test_point_direction = ttk.Combobox(dut_test_step_direction_frame,
                                            values=test_point_direction_list)
        acc.always_active_style(test_point_direction)
        test_point_direction.config(state="active", width=error_type_width)
        test_point_direction.grid(pady=5, row=1, column=2, columnspan=1)

        # These labels are dummy labels that do nothing but help format the frame
        dummysd = ttk.Label(dut_test_step_direction_frame)
        dummysd.grid(row=1, column=3)
        dummysd.config(width=2)
        dummysd_1 = ttk.Label(dut_test_step_direction_frame)
        dummysd_1.grid(row=1, column=0)
        dummysd_1.config(width=2)

        if dut_output_type.get() == "Dual":
            # Test Step Number Designation
            test_point_number_list_1 = [" ", "1", "2", "3", "4", "5", "6", "7", "8", "9"]
            device_under_test_number_of_test_points_1 = StringVar(DUTCalDetails)
            lbl_number_of_test_points_1 = ttk.Label(dut_test_step_direction_frame_1,
                                                    text="Number of Test Points:", width=20, font=('arial', 12),
                                                    anchor="n")
            lbl_number_of_test_points_1.grid(row=0, column=1)
            number_of_test_points_1 = ttk.Combobox(dut_test_step_direction_frame_1,
                                                   values=test_point_number_list_1)
            acc.always_active_style(number_of_test_points_1)
            number_of_test_points_1.config(state="active", width=error_type_width)
            number_of_test_points_1.grid(pady=5, row=0, column=2, columnspan=1)

            # Test Point Direction Assignment
            test_point_direction_list_1 = [" ", "Ascending", "Descending", "Asc. & Desc."]
            device_under_test_test_point_direction_1 = StringVar(DUTCalDetails)
            lbl_test_point_direction = ttk.Label(dut_test_step_direction_frame_1, text="Direction:", width=20,
                                                 font=('arial', 12), anchor="n")
            lbl_test_point_direction.grid(row=1, column=1)
            test_point_direction_1 = ttk.Combobox(dut_test_step_direction_frame_1,
                                                  values=test_point_direction_list_1)
            acc.always_active_style(test_point_direction_1)
            test_point_direction_1.config(state="active", width=error_type_width)
            test_point_direction_1.grid(pady=5, row=1, column=2, columnspan=1)

            # These labels are dummy labels that do nothing but help format the frame
            dummysd1 = ttk.Label(dut_test_step_direction_frame_1)
            dummysd1.grid(row=1, column=3)
            dummysd1.config(width=2)
            dummysd1_1 = ttk.Label(dut_test_step_direction_frame_1)
            dummysd1_1.grid(row=1, column=0)
            dummysd1_1.config(width=2)

        # .................Accuracy/Specification Calculation....................#

        # Frame Creation for Accuracy/Specification of Unit
        if dut_output_type.get() == "Single":
            device_accuracy_frame = LabelFrame(tab1, text="DUT Accuracy/Specification", relief=SOLID, bd=1,
                                               labelanchor="n")
            device_accuracy_frame.grid(pady=5, row=8, column=1, rowspan=2, columnspan=6)
            dummyasf = ttk.Label(tab1)
            dummyasf.grid(row=8, column=0)
            dummyasf.config(width=2)
            dummyasf_1 = ttk.Label(tab1)
            dummyasf_1.grid(row=8, column=8)
            dummyasf_1.config(width=2)
        elif dut_output_type.get() == "Dual":
            device_accuracy_frame = LabelFrame(tab2, text="Discipline 1: DUT Accuracy/Specification",
                                               relief=SOLID, bd=1, labelanchor="n")
            device_accuracy_frame.grid(pady=5, row=8, column=1, rowspan=2, columnspan=3)
            device_accuracy_frame_1 = LabelFrame(tab2, text="Discipline 2: DUT Accuracy/Specification",
                                                 relief=SOLID, bd=1, labelanchor="n")
            device_accuracy_frame_1.grid(pady=5, row=8, column=5, rowspan=2, columnspan=3)
            dummyasfdual = ttk.Label(tab2)
            dummyasfdual.grid(row=8, column=0)
            dummyasfdual.config(width=1)
            dummyasfdual_1 = ttk.Label(tab2)
            dummyasfdual_1.grid(row=8, column=4)
            dummyasfdual_1.config(width=1)
            dummyasfdual_2 = ttk.Label(tab2)
            dummyasfdual_2.grid(row=8, column=9)
            dummyasfdual_2.config(width=1)
        elif dut_output_type.get() == "Transmitter":
            device_accuracy_frame = LabelFrame(tab3, text="DUT Accuracy/Specification", relief=SOLID, bd=1,
                                               labelanchor="n")
            device_accuracy_frame.grid(pady=5, row=10, column=1, rowspan=2, columnspan=6)
            dummyasftrans = ttk.Label(tab3)
            dummyasftrans.grid(row=10, column=0)
            dummyasftrans.config(width=2)
            dummyasftrans_1 = ttk.Label(tab3)
            dummyasftrans_1.grid(row=10, column=8)
            dummyasftrans_1.config(width=2)

        # Operator Designation
        # Less-Than or Equal To: 2264
        # Greater-Than or Equal To: 2265
        # Less-Than: 02C2
        # Greater-Than: 02C3
        # Plus-Minus Sign: 00B1
        first_operator_list = [" ", "<", ">", u"\u2264", u"\u2265", u"\u00B1"]
        supplemental_operator_list = ["", "+", "-", u"\u00B1"]
        selected_operator = StringVar(DUTCalDetails)
        selected_operator_1 = StringVar(DUTCalDetails)
        selected_operator_2 = StringVar(DUTCalDetails)
        selected_operator_3 = StringVar(DUTCalDetails)

        # Error Assignment
        measurement_error_assignment_type = StringVar(DUTCalDetails)
        measurement_error_assignment_type_1 = StringVar(DUTCalDetails)
        measurement_error_assignment_type_2 = StringVar(DUTCalDetails)
        measurement_error_assignment_type_3 = StringVar(DUTCalDetails)

        # Width Definition for Operators
        operator_width = len(max(supplemental_operator_list, key=len))

        # Possible Accuracy #1
        operator_drop_down = ttk.Combobox(device_accuracy_frame,
                                          values=first_operator_list)
        acc.always_active_style(operator_drop_down)
        operator_drop_down.config(state="active", width=operator_width)
        operator_drop_down.grid(pady=5, padx=2, row=0, column=1)
        measurement_error_type_1_value = ttk.Entry(device_accuracy_frame,
                                                   textvariable=device_under_test_error_selection_value,
                                                   font=('arial', 12))
        measurement_error_type_1_value.grid(row=0, column=2)
        measurement_error_type_1_value.config(width=14)
        device_under_test_error_selection = ttk.Combobox(device_accuracy_frame,
                                                         values=error_type_drop_down_list)
        acc.always_active_style(device_under_test_error_selection)
        device_under_test_error_selection.config(state="active", width=error_type_width)
        device_under_test_error_selection.grid(pady=5, padx=2, row=0, column=3)

        # Possible Accuracy #2
        operator_drop_down_1 = ttk.Combobox(device_accuracy_frame, values=supplemental_operator_list)
        acc.always_active_style(operator_drop_down_1)
        operator_drop_down_1.config(state="active", width=operator_width)
        operator_drop_down_1.grid(pady=5, padx=2, row=0, column=4)
        measurement_error_type_1_value_additional = ttk.Entry(device_accuracy_frame,
                                                              textvariable=device_under_test_error_selection_value_1,
                                                              font=('arial', 12))
        measurement_error_type_1_value_additional.grid(row=0, column=5)
        measurement_error_type_1_value_additional.config(width=14)
        device_under_test_error_selection_1 = ttk.Combobox(device_accuracy_frame,
                                                           values=error_type_drop_down_list)
        acc.always_active_style(device_under_test_error_selection_1)
        device_under_test_error_selection_1.config(state="active", width=error_type_width)
        device_under_test_error_selection_1.grid(pady=5, padx=2, row=0, column=6)

        # Possible Accuracy #3
        operator_drop_down_2 = ttk.Combobox(device_accuracy_frame, values=supplemental_operator_list)
        acc.always_active_style(operator_drop_down_2)
        operator_drop_down_2.config(state="active", width=operator_width)
        operator_drop_down_2.grid(pady=5, padx=2, row=1, column=1)
        measurement_error_type_1_value_additional_1 = ttk.Entry(device_accuracy_frame,
                                                                textvariable=device_under_test_error_selection_value_2,
                                                                font=('arial', 12))
        measurement_error_type_1_value_additional_1.grid(row=1, column=2)
        measurement_error_type_1_value_additional_1.config(width=14)
        device_under_test_error_selection_2 = ttk.Combobox(device_accuracy_frame, values=error_type_drop_down_list)
        acc.always_active_style(device_under_test_error_selection_2)
        device_under_test_error_selection_2.config(state="active", width=error_type_width)
        device_under_test_error_selection_2.grid(pady=5, padx=2, row=1, column=3)

        # Possible Accuracy #4
        operator_drop_down_3 = ttk.Combobox(device_accuracy_frame, values=supplemental_operator_list)
        acc.always_active_style(operator_drop_down_3)
        operator_drop_down_3.config(state="active", width=operator_width)
        operator_drop_down_3.grid(pady=5, padx=2, row=1, column=4)
        measurement_error_type_1_value_additional_2 = ttk.Entry(device_accuracy_frame,
                                                                textvariable=device_under_test_error_selection_value_3,
                                                                font=('arial', 12))
        measurement_error_type_1_value_additional_2.grid(row=1, column=5)
        measurement_error_type_1_value_additional_2.config(width=14)
        device_under_test_error_selection_3 = ttk.Combobox(device_accuracy_frame, values=error_type_drop_down_list)
        acc.always_active_style(device_under_test_error_selection_3)
        device_under_test_error_selection_3.config(state="active", width=error_type_width)
        device_under_test_error_selection_3.grid(pady=5, padx=2, row=1, column=6)

        # These labels are dummy labels that do nothing but help format the frame
        dummyas = ttk.Label(device_accuracy_frame)
        dummyas.grid(row=1, column=0)
        dummyas.config(width=2)
        dummyas_1 = ttk.Label(device_accuracy_frame)
        dummyas_1.grid(row=1, column=7)
        dummyas_1.config(width=2)

        if dut_output_type.get() == "Dual":
            # Operator Designation
            # Less-Than or Equal To: 2264
            # Greater-Than or Equal To: 2265
            # Less-Than: 02C2
            # Greater-Than: 02C3
            # Plus-Minus Sign: 00B1
            first_operator_list = [" ", "<", ">", u"\u2264", u"\u2265", u"\u00B1"]
            supplemental_operator_list = ["", "+", "-", u"\u00B1"]
            second_selected_operator = StringVar(DUTCalDetails)
            second_selected_operator_1 = StringVar(DUTCalDetails)
            second_selected_operator_2 = StringVar(DUTCalDetails)
            second_selected_operator_3 = StringVar(DUTCalDetails)

            # Error Assignment
            second_measurement_error_assignment_type = StringVar(DUTCalDetails)
            second_measurement_error_assignment_type_1 = StringVar(DUTCalDetails)
            second_measurement_error_assignment_type_2 = StringVar(DUTCalDetails)
            second_measurement_error_assignment_type_3 = StringVar(DUTCalDetails)

            # Possible Accuracy #1
            second_operator_drop_down = ttk.Combobox(device_accuracy_frame_1, values=first_operator_list)
            acc.always_active_style(second_operator_drop_down)
            second_operator_drop_down.config(state="active", width=operator_width)
            second_operator_drop_down.grid(pady=5, padx=2, row=0, column=1)
            second_measurement_error_type_value = ttk.Entry(device_accuracy_frame_1,
                                                            textvariable=device_under_test_second_error_selection_value,
                                                            font=('arial', 12))
            second_measurement_error_type_value.grid(row=0, column=2)
            second_measurement_error_type_value.config(width=14)
            second_device_under_test_error_selection = ttk.Combobox(device_accuracy_frame_1,
                                                                    values=error_type_drop_down_list)
            acc.always_active_style(second_device_under_test_error_selection)
            second_device_under_test_error_selection.config(state="active", width=error_type_width)
            second_device_under_test_error_selection.grid(pady=5, padx=2, row=0, column=3)

            # Possible Accuracy #2
            second_operator_drop_down_1 = ttk.Combobox(device_accuracy_frame_1, values=supplemental_operator_list)
            acc.always_active_style(second_operator_drop_down_1)
            second_operator_drop_down_1.config(state="active", width=operator_width)
            second_operator_drop_down_1.grid(pady=5, padx=2, row=0, column=4)
            second_measurement_error_type_value_1 = ttk.Entry(device_accuracy_frame_1,
                                                              textvariable=device_under_test_second_error_selection_value_1,
                                                              font=('arial', 12))
            second_measurement_error_type_value_1.grid(row=0, column=5)
            second_measurement_error_type_value_1.config(width=14)
            second_device_under_test_error_selection_1 = ttk.Combobox(device_accuracy_frame_1,
                                                                      values=error_type_drop_down_list)
            acc.always_active_style(second_device_under_test_error_selection_1)
            second_device_under_test_error_selection_1.config(state="active", width=error_type_width)
            second_device_under_test_error_selection_1.grid(pady=5, padx=2, row=0, column=6)

            # Possible Accuracy #3
            second_operator_drop_down_2 = ttk.Combobox(device_accuracy_frame_1, values=supplemental_operator_list)
            acc.always_active_style(second_operator_drop_down_2)
            second_operator_drop_down_2.config(state="active", width=operator_width)
            second_operator_drop_down_2.grid(pady=5, padx=2, row=1, column=1)
            second_measurement_error_type_value_2 = ttk.Entry(device_accuracy_frame_1,
                                                              textvariable=device_under_test_second_error_selection_value_2,
                                                              font=('arial', 12))
            second_measurement_error_type_value_2.grid(row=1, column=2)
            second_measurement_error_type_value_2.config(width=14)
            second_device_under_test_error_selection_2 = ttk.Combobox(device_accuracy_frame_1,
                                                                      values=error_type_drop_down_list)
            acc.always_active_style(second_device_under_test_error_selection_2)
            second_device_under_test_error_selection_2.config(state="active", width=error_type_width)
            second_device_under_test_error_selection_2.grid(pady=5, padx=2, row=1, column=3)

            # Possible Accuracy #4
            second_operator_drop_down_3 = ttk.Combobox(device_accuracy_frame_1, values=supplemental_operator_list)
            acc.always_active_style(second_operator_drop_down_3)
            second_operator_drop_down_3.config(state="active", width=operator_width)
            second_operator_drop_down_3.grid(pady=5, padx=2, row=1, column=4)
            second_measurement_error_type_value_3 = ttk.Entry(device_accuracy_frame_1,
                                                              textvariable=device_under_test_second_error_selection_value_3,
                                                              font=('arial', 12))
            second_measurement_error_type_value_3.grid(row=1, column=5)
            second_measurement_error_type_value_3.config(width=14)
            second_device_under_test_error_selection_3 = ttk.Combobox(device_accuracy_frame_1,
                                                                      values=error_type_drop_down_list)
            acc.always_active_style(second_device_under_test_error_selection_3)
            second_device_under_test_error_selection_3.config(state="active", width=error_type_width)
            second_device_under_test_error_selection_3.grid(pady=5, padx=2, row=1, column=6)

            # These labels are dummy labels that do nothing but help format the frame
            dummyas1 = ttk.Label(device_accuracy_frame_1)
            dummyas1.grid(row=1, column=0)
            dummyas1.config(width=2)
            dummyas1_1 = ttk.Label(device_accuracy_frame_1)
            dummyas1_1.grid(row=1, column=7)
            dummyas1_1.config(width=2)

        # .....................Button Labels and Buttons.........................#

        if dut_output_type.get() == "Single" or \
                dut_output_type.get() == "Transmitter":
            btn_back_out_calibration_specifications = ttk.Button(DUTCalDetails, text="Back", width=20,
                                                                 command=lambda: self.device_instrument_description_information(DUTCalDetails))
            btn_back_out_calibration_specifications.bind("<Return>", lambda event: self.device_instrument_description_information(DUTCalDetails))
            btn_back_out_calibration_specifications.grid(pady=5, row=3, column=1, columnspan=1)

            btn_advance_to_reference_input = ttk.Button(DUTCalDetails, text="Next", width=20,
                                                        command=lambda: self.verify_applied_dut_calibration_specifications())
            btn_advance_to_reference_input.bind("<Return>", lambda event: self.verify_applied_dut_calibration_specifications())
            btn_advance_to_reference_input.grid(pady=5, row=3, column=5, columnspan=1)
        else:
            btn_back_out_calibration_specifications = ttk.Button(DUTCalDetails, text="Back", width=20,
                                                                 command=lambda: self.device_instrument_description_information(DUTCalDetails))
            btn_back_out_calibration_specifications.bind("<Return>", lambda event: self.device_instrument_description_information(DUTCalDetails))
            btn_back_out_calibration_specifications.grid(pady=5, row=3, column=2, columnspan=1)

            btn_advance_to_reference_input = ttk.Button(DUTCalDetails, text="Next", width=20,
                                                        command=lambda: self.verify_applied_dut_calibration_specifications())
            btn_advance_to_reference_input.bind("<Return>", lambda event: self.verify_applied_dut_calibration_specifications())
            btn_advance_to_reference_input.grid(pady=5, row=3, column=4, columnspan=1)

    # -----------------------------------------------------------------------#

    # Command to Change Layout of DUTCalDetails Depending On User Selection of Output of DUT.
    def update_dut_measurement_unit_selection(self):
        global units_of_measure_list, device_under_test_units, units_of_measure_list_1, device_under_test_units_1

        # ..................Update Discipline 1 Units Drop down...................#

        if measurement_discipline_selection.get() == "Capacitance":
            units_of_measure_list = [" ", "pF", "nF", u"\u03BC" + "F", "mF", "F"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Current":
            units_of_measure_list = [" ", "nA AC", u"\u03BC" + "A AC", "mA AC", "A AC", "nA DC", u"\u03BC" + "A DC",
                                     "mA DC", "A DC"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Concentration":
            units_of_measure_list = [" ", "mol/L", "ppb", "ppm", "ppth"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Flow":
            units_of_measure_list = [" ", "CCM", "CFM", "ft/s", "GPH", "GPM", "l/s", "l/h",
                                     "LPM", "CMH", "CMS", "ml/min", "SCFH", "SCFM"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Frequency":
            units_of_measure_list = [" ", "mHz", "Hz", "kHz", "MHz", "GHz"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Humidity":
            units_of_measure_list = [" ", "%RH"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Pressure":
            units_of_measure_list = [" ", "mbar", "bar", "Pa", "hPa", "kPa", "MPa", "ksi", "psia", "psig", "psid",
                                     "kg/ccm", "oz/in", "cm W.C.", "ft W.C.", "in W.C.",
                                     "in W.C. at 60" + u"\u00B0" + "F", "in W.C. at 20" + u"\u00B0" + "C", "mm W.C.",
                                     "in Hg", "mm Hg", "Torr"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Resistance":
            units_of_measure_list = [" ", u"\u03BC" + u"\u03A9", "m" + u"\u03A9", u"\u03A9", "k" + u"\u03A9",
                                     "M" + u"\u03A9", "G" + u"\u03A9"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Temperature":
            units_of_measure_list = [" ", u"\u00B0" + "C", u"\u00B0" + "F", u"\u00B0" + "K"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Velocity":
            units_of_measure_list = [" ", "m/s", "FPM"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        elif measurement_discipline_selection.get() == "Voltage":
            units_of_measure_list = [" ", "pV AC", "nV AC", u"\u03BC" + "V AC", "mV AC", "V AC", "kV AC", "pV DC",
                                     "nV DC", u"\u03BC" + "V DC", "mV DC", "V DC", "kV DC"]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        else:
            units_of_measure_list = [" "]
            device_under_test_units.config(state="active", width=error_type_width, values=units_of_measure_list)
            device_under_test_units.grid(pady=5, row=0, column=5, columnspan=1)

        # ..................Update Measurement 2 Units Drop down...................#

        if dut_output_type.get() == "Dual":
            if measurement_discipline_selection_1.get() == "Capacitance":
                units_of_measure_list_1 = [" ", "pF", "nF", u"\u03BC" + "F", "mF", "F"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Current":
                units_of_measure_list_1 = [" ", "nA AC", u"\u03BC" + "A AC", "mA AC", "A AC", "nA DC",
                                           u"\u03BC" + "A DC", "mA DC", "A DC"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Concentration":
                units_of_measure_list_1 = [" ", "mol/L", "ppb", "ppm", "ppth"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Flow":
                units_of_measure_list_1 = [" ", "CCM", "CFM", "ft/s", "GPH", "GPM", "l/s", "l/h",
                                           "LPM", "CMH", "CMS", "ml/min", "SCFH", "SCFM"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Frequency":
                units_of_measure_list_1 = [" ", "mHz", "Hz", "kHz", "MHz", "GHz"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Humidity":
                units_of_measure_list_1 = [" ", "%RH"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Pressure":
                units_of_measure_list_1 = [" ", "mbar", "bar", "Pa", "hPa", "kPa", "MPa", "ksi", "psia", "psig",
                                           "psid", "kg/ccm", "oz/in", "cm W.C.", "ft W.C.", "in W.C.",
                                           "in W.C. at 60" + u"\u00B0" + "F", "in W.C. at 20" + u"\u00B0" + "C",
                                           "mm W.C.", "in Hg", "mm Hg", "Torr"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Resistance":
                units_of_measure_list_1 = [" ", u"\u03BC" + u"\u03A9", "m" + u"\u03A9", u"\u03A9", "k" + u"\u03A9",
                                           "M" + u"\u03A9", "G" + u"\u03A9"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Temperature":
                units_of_measure_list_1 = [" ", u"\u00B0" + "C", u"\u00B0" + "F", u"\u00B0" + "K"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Velocity":
                units_of_measure_list_1 = [" ", "m/s", "FPM"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Voltage":
                units_of_measure_list_1 = [" ", "pV AC", "nV AC", u"\u03BC" + "V AC", "mV AC", "V AC", "kV AC", "pV DC",
                                           "nV DC", u"\u03BC" + "V DC", "mV DC", "V DC", "kV DC"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            else:
                units_of_measure_list_1 = [" "]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

        elif dut_output_type.get() == "Transmitter":
            if measurement_discipline_selection_1.get() == "Voltage":
                units_of_measure_list_1 = [" ", "mV", "V", "kV"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            elif measurement_discipline_selection_1.get() == "Current":
                units_of_measure_list_1 = [" ", "mA", "A"]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

            else:
                units_of_measure_list_1 = [" "]
                device_under_test_units_1.config(state="active", width=error_type_width,
                                                 values=units_of_measure_list_1)
                device_under_test_units_1.grid(pady=5, row=1, column=5, columnspan=1)

    # -----------------------------------------------------------------------#

    # Command to Save All Fields Set By User for Certificate of Calibration
    def apply_dut_calibration_specifications(self):

        global condition_of_dut, device_under_test_output_type, device_under_test_measurement_type, units_of_measure, \
            device_under_test_minimum, device_under_test_maximum, reference_resolution, \
            device_under_test_measurement_resolution, dut_number_of_test_points, \
            device_under_test_calibration_direction, device_under_test_specification_operator, \
            device_under_test_specification_operator_1, device_under_test_specification_operator_2, \
            device_under_test_specification_operator_3, device_under_test_specification_value, \
            device_under_test_specification_value_1, device_under_test_specification_value_2, \
            device_under_test_specification_value_3, device_under_test_specification_type, \
            device_under_test_specification_type_1, device_under_test_specification_type_2, \
            device_under_test_specification_type_3, second_device_under_test_measurement_type, \
            second_units_of_measure, second_device_under_test_minimum, second_device_under_test_maximum, \
            second_reference_resolution, second_device_under_test_measurement_resolution, \
            second_device_under_test_number_of_test_points, second_device_under_test_calibration_direction, \
            second_device_under_test_specification_operator, second_device_under_test_specification_operator_1, \
            second_device_under_test_specification_operator_2, second_device_under_test_specification_operator_3, \
            second_device_under_test_specification_value, second_device_under_test_specification_value_1, \
            second_device_under_test_specification_value_2, second_device_under_test_specification_value_3, \
            second_device_under_test_specification_type, second_device_under_test_specification_type_1, \
            second_device_under_test_specification_type_2, second_device_under_test_specification_type_3

        # Save Values Applied on Calibration Parameters Window

        condition_of_dut = dut_condition.get()
        device_under_test_output_type = dut_output_type.get()

        # .............Save Measurement 1 Information to Variables...............#

        device_under_test_measurement_type = measurement_discipline_selection.get()
        units_of_measure = device_under_test_units.get()
        device_under_test_minimum = device_under_test_minimum_full_scale.get()
        device_under_test_maximum = device_under_test_maximum_full_scale.get()
        reference_resolution = reference_resolution_value.get()
        device_under_test_measurement_resolution = device_under_test_resolution.get()
        dut_number_of_test_points = number_of_test_points.get()
        device_under_test_calibration_direction = test_point_direction.get()
        device_under_test_specification_operator = operator_drop_down.get()
        device_under_test_specification_operator_1 = operator_drop_down_1.get()
        device_under_test_specification_operator_2 = operator_drop_down_2.get()
        device_under_test_specification_operator_3 = operator_drop_down_3.get()
        device_under_test_specification_value = measurement_error_type_1_value.get()
        device_under_test_specification_value_1 = measurement_error_type_1_value_additional.get()
        device_under_test_specification_value_2 = measurement_error_type_1_value_additional_1.get()
        device_under_test_specification_value_3 = measurement_error_type_1_value_additional_2.get()
        device_under_test_specification_type = device_under_test_error_selection.get()
        device_under_test_specification_type_1 = device_under_test_error_selection_1.get()
        device_under_test_specification_type_2 = device_under_test_error_selection_2.get()
        device_under_test_specification_type_3 = device_under_test_error_selection_3.get()

        LIMSVarConfig.imported_full_scale = device_under_test_maximum
        LIMSVarConfig.imported_specification_value = device_under_test_specification_value

        # .............Save Measurement 2 Information to Variables...............#

        if dut_output_type.get() == "Dual" or dut_output_type.get() == "Transmitter":
            second_device_under_test_measurement_type = measurement_discipline_selection_1.get()
            second_units_of_measure = device_under_test_units_1.get()
            second_device_under_test_minimum = device_under_test_minimum_full_scale_1.get()
            second_device_under_test_maximum = device_under_test_maximum_full_scale_1.get()
            if dut_output_type.get() == "Dual":
                second_reference_resolution = reference_resolution_value_1.get()
                second_device_under_test_measurement_resolution = device_under_test_resolution_1.get()
                second_device_under_test_number_of_test_points = number_of_test_points_1.get()
                second_device_under_test_calibration_direction = test_point_direction_1.get()
                second_device_under_test_specification_operator = second_operator_drop_down.get()
                second_device_under_test_specification_operator_1 = second_operator_drop_down_1.get()
                second_device_under_test_specification_operator_2 = second_operator_drop_down_2.get()
                second_device_under_test_specification_operator_3 = second_operator_drop_down_3.get()
                second_device_under_test_specification_value = second_measurement_error_type_value.get()
                second_device_under_test_specification_value_1 = second_measurement_error_type_value_1.get()
                second_device_under_test_specification_value_2 = second_measurement_error_type_value_2.get()
                second_device_under_test_specification_value_3 = second_measurement_error_type_value_3.get()
                second_device_under_test_specification_type = second_device_under_test_error_selection.get()
                second_device_under_test_specification_type_1 = second_device_under_test_error_selection_1.get()
                second_device_under_test_specification_type_2 = second_device_under_test_error_selection_2.get()
                second_device_under_test_specification_type_3 = second_device_under_test_error_selection_3.get()
            else:
                pass

    # -----------------------------------------------------------------------#

    # Command to Verify All User Inputs Applied Prior to Further Calibration Process
    def verify_applied_dut_calibration_specifications(self):

        self.apply_dut_calibration_specifications()

        if dut_output_type.get() == "Single":
            if measurement_discipline_selection.get() == "" or measurement_discipline_selection.get() == " " or \
                    device_under_test_units.get() == "" or device_under_test_units.get() == " " or \
                    device_under_test_minimum_full_scale.get() == "" or \
                    device_under_test_maximum_full_scale.get() == "" or \
                    reference_resolution_value.get() == "" or \
                    device_under_test_resolution.get() == "" or \
                    number_of_test_points.get() == "" or number_of_test_points.get() == " " or \
                    test_point_direction.get() == "" or test_point_direction.get() == " " or \
                    operator_drop_down.get() == "" or operator_drop_down == " " or \
                    measurement_error_type_1_value.get() == "" or device_under_test_error_selection.get() == "" or \
                    device_under_test_error_selection.get() == " ":
                tm.showerror("Error", 'Please complete the required fields!')
            else:
                self.calibration_standards_selection(DUTCalDetails)
        elif dut_output_type.get() == "Transmitter":
            if measurement_discipline_selection.get() == "" or measurement_discipline_selection.get() == " " or \
                    device_under_test_units.get() == "" or device_under_test_units.get() == " " or \
                    device_under_test_minimum_full_scale.get() == "" or \
                    device_under_test_maximum_full_scale.get() == "" or \
                    reference_resolution_value.get() == "" or \
                    device_under_test_resolution.get() == "" or \
                    number_of_test_points.get() == "" or number_of_test_points.get() == " " or \
                    test_point_direction.get() == "" or test_point_direction.get() == " " or \
                    operator_drop_down.get() == "" or operator_drop_down.get() == " " or \
                    measurement_error_type_1_value.get() == "" or \
                    device_under_test_error_selection.get() == "" or device_under_test_error_selection.get() == " " or \
                    measurement_discipline_selection_1.get() == " " or \
                    device_under_test_units_1.get() == "" or device_under_test_units_1.get() == " " or \
                    device_under_test_minimum_full_scale_1.get() == " " or \
                    device_under_test_maximum_full_scale_1.get() == " ":
                tm.showerror("Error", 'Please complete the required fields!')
            else:
                self.calibration_standards_selection(DUTCalDetails)
        elif dut_output_type.get() == "Dual":
            if measurement_discipline_selection.get() == "" or measurement_discipline_selection.get() == " " or \
                    measurement_discipline_selection_1.get() == "" or \
                    measurement_discipline_selection_1.get() == " " or \
                    device_under_test_units.get() == "" or device_under_test_units.get() == " " or \
                    device_under_test_units_1.get() == "" or device_under_test_units_1.get() == " " or \
                    device_under_test_minimum_full_scale.get() == "" or \
                    device_under_test_minimum_full_scale_1.get() == "" or \
                    device_under_test_maximum_full_scale.get() == "" or \
                    device_under_test_maximum_full_scale_1.get() == "" or \
                    reference_resolution_value.get() == "" or reference_resolution_value_1.get() == "" or \
                    device_under_test_resolution.get() == "" or device_under_test_resolution_1.get() == "" or \
                    number_of_test_points.get() == "" or number_of_test_points.get() == " " or \
                    number_of_test_points_1.get() == "" or number_of_test_points_1.get() == " " or \
                    test_point_direction.get() == "" or test_point_direction.get() == " " or \
                    test_point_direction_1.get() == "" or test_point_direction_1.get() == " " or \
                    operator_drop_down.get() == "" or operator_drop_down.get() == " " or \
                    second_operator_drop_down.get() == "" or second_operator_drop_down.get() == " " or \
                    measurement_error_type_1_value.get() == "" or second_measurement_error_type_value.get() == "" or \
                    device_under_test_error_selection.get() == "" or device_under_test_error_selection.get() == " " or \
                    second_device_under_test_error_selection.get() == "" or \
                    second_device_under_test_error_selection.get() == " ":
                tm.showerror("Error", 'Please complete the required fields!')
            else:
                self.calibration_standards_selection(DUTCalDetails)

    # -----------------------------------------------------------------------#

    # Command to Send User to Load Calibration Standards Section of Calibration Module
    def calibration_standards_selection(self, window):
        self.window = window
        window.withdraw()

        LIMSVarConfig.clear_loaded_calibration_standard_information_variables()

        global CalibrationStandardDetails, calibration_standard_1, calibration_standard_2, calibration_standard_3, \
            calibration_standard_4, calibration_standard_5, calibration_standard_6, calibration_standard_7, \
            calibration_standard_8, standard_description, standard_description_2, standard_description_3, \
            standard_description_4, standard_description_5, standard_description_6, standard_description_7, \
            standard_description_8, standard_serial_number, standard_serial_number_2, standard_serial_number_3, \
            standard_serial_number_4, standard_serial_number_5, standard_serial_number_6, standard_serial_number_7, \
            standard_serial_number_8, standard_calibration_date, standard_calibration_date_2, \
            standard_calibration_date_3, standard_calibration_date_4, standard_calibration_date_5, \
            standard_calibration_date_6, standard_calibration_date_7, standard_calibration_date_8, \
            standard_calibration_due_date, standard_calibration_due_date_2, standard_calibration_due_date_3, \
            standard_calibration_due_date_4, standard_calibration_due_date_5, standard_calibration_due_date_6, \
            standard_calibration_due_date_7, standard_calibration_due_date_8

        calibration_standard_1_variable = StringVar()
        calibration_standard_2_variable = StringVar()
        calibration_standard_3_variable = StringVar()
        calibration_standard_4_variable = StringVar()
        calibration_standard_5_variable = StringVar()
        calibration_standard_6_variable = StringVar()
        calibration_standard_7_variable = StringVar()
        calibration_standard_8_variable = StringVar()

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        aph = AppHelpWindows()

        # ...................Reference Calibration Standards.....................#

        CalibrationStandardDetails = Toplevel()
        CalibrationStandardDetails.title("Calibration Reference Standards")
        CalibrationStandardDetails.iconbitmap("Required Images\\DwyerLogo.ico")
        height = 420
        width = 1360
        screen_width = CalibrationStandardDetails.winfo_screenwidth()
        screen_height = CalibrationStandardDetails.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        CalibrationStandardDetails.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CalibrationStandardDetails.focus_force()
        CalibrationStandardDetails.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CalibrationStandardDetails))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(CalibrationStandardDetails, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CalibrationStandardDetails)),
                                           ("DUT Calibration Information",
                                            lambda: self.device_under_test_calibration_specifications(CalibrationStandardDetails)),
                                           ("Logout", lambda: acc.software_signout(CalibrationStandardDetails)),
                                           ("Quit", lambda: acc.software_close(CalibrationStandardDetails))])
        menubar.add_menu("Help", commands=[("Help", lambda: aph.calibration_standard_loading_details())])

        # ...........................Frame Creation...............................#

        # Frame Creation for Calibration Standards
        calibration_standard_frame = LabelFrame(CalibrationStandardDetails,
                                                text="List the Calibration Standards Used During Calibration",
                                                relief=SOLID, bd=1, labelanchor="n")
        calibration_standard_frame.grid(padx=10, pady=5, row=0, column=0, rowspan=7, columnspan=6)

        # Width Definitions for Calibration Standards
        description_width = len(max(LIMSVarConfig.cal_equip_descrip, key=len))
        serial_number_width = len(max(LIMSVarConfig.cal_equip_serial_no, key=len))
        date_of_calibration_width = len(max(LIMSVarConfig.cal_equip_cal_date, key=len))
        calibration_due_date_width = len(max(LIMSVarConfig.cal_equip_due_date, key=len))

        # Headers for Reference Calibration Standards
        lbl_description = ttk.Label(calibration_standard_frame, text="Description", font=('arial', 11), anchor="n")
        lbl_description.grid(row=0, column=2, pady=5)
        lbl_serial_number = ttk.Label(calibration_standard_frame, text="Serial Number", font=('arial', 11), anchor="n")
        lbl_serial_number.grid(row=0, column=3)
        lbl_calibration_date = ttk.Label(calibration_standard_frame, text="Cal. Date", font=('arial', 11), anchor="n")
        lbl_calibration_date.grid(row=0, column=4)
        lbl_due_date = ttk.Label(calibration_standard_frame, text="Due Date", font=('arial', 11), anchor="n")
        lbl_due_date.grid(row=0, column=5)

        # Asset Number 1
        lbl_asset_number = ttk.Label(calibration_standard_frame, text="1) Asset Number:", font=('arial', 12),
                                     anchor="n")
        lbl_asset_number.grid(row=1, column=0, pady=5, padx=5)
        calibration_standard_1 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_1_variable,
                                           font=('arial', 12))
        calibration_standard_1.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_1_variable))
        calibration_standard_1.grid(row=1, column=1)
        calibration_standard_1.config(width=15)
        calibration_standard_1.focus()

        # Asset Number 1 - Description, Serial, Cal, Due Date
        standard_description = ttk.Label(calibration_standard_frame,
                                         text=LIMSVarConfig.calibration_standard_equipment_description,
                                         font=10, anchor="n")
        standard_description.grid(row=1, column=2)
        standard_description.config(width=description_width)
        standard_serial_number = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_serial_number, font=10, anchor="n")
        standard_serial_number.grid(row=1, column=3)
        standard_serial_number.config(width=serial_number_width)
        standard_calibration_date = ttk.Label(calibration_standard_frame,
                                              text=LIMSVarConfig.calibration_standard_equipment_calibration_date,
                                              font=10, anchor="n")
        standard_calibration_date.grid(row=1, column=4)
        standard_calibration_date.config(width=date_of_calibration_width)
        standard_calibration_due_date = ttk.Label(calibration_standard_frame,
                                                  text=LIMSVarConfig.calibration_standard_equipment_due_date, font=10,
                                                  anchor="n")
        standard_calibration_due_date.grid(row=1, column=5)
        standard_calibration_due_date.config(width=calibration_due_date_width)

        # Asset Number 2
        lbl_asset_number_2 = ttk.Label(calibration_standard_frame, text="2) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_2.grid(row=2, column=0, pady=5, padx=5)
        calibration_standard_2 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_2_variable,
                                           font=('arial', 12))
        calibration_standard_2.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_2_variable))
        calibration_standard_2.grid(row=2, column=1)
        calibration_standard_2.config(width=15)

        # Asset Number 2 - Description, Serial, Cal, Due Date
        standard_description_2 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_2, font=10,
                                           anchor="n")
        standard_description_2.grid(row=2, column=2)
        standard_description_2.config(width=description_width)
        standard_serial_number_2 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_2, font=10,
                                             anchor="n")
        standard_serial_number_2.grid(row=2, column=3)
        standard_serial_number_2.config(width=serial_number_width)
        standard_calibration_date_2 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_2,
                                                font=10, anchor="n")
        standard_calibration_date_2.grid(row=2, column=4)
        standard_calibration_date_2.config(width=date_of_calibration_width)
        standard_calibration_due_date_2 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_2,
                                                    font=10, anchor="n")
        standard_calibration_due_date_2.grid(row=2, column=5)
        standard_calibration_due_date_2.config(width=calibration_due_date_width)

        # Asset Number 3
        lbl_asset_number_3 = ttk.Label(calibration_standard_frame, text="3) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_3.grid(row=3, column=0, pady=5, padx=5)
        calibration_standard_3 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_3_variable,
                                           font=('arial', 12))
        calibration_standard_3.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_3_variable))
        calibration_standard_3.grid(row=3, column=1)
        calibration_standard_3.config(width=15)

        # Asset Number 3 - Description, Serial, Cal, Due Date
        standard_description_3 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_3, font=10,
                                           anchor="n")
        standard_description_3.grid(row=3, column=2)
        standard_description_3.config(width=description_width)
        standard_serial_number_3 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_3, font=10,
                                             anchor="n")
        standard_serial_number_3.grid(row=3, column=3)
        standard_serial_number_3.config(width=serial_number_width)
        standard_calibration_date_3 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_3,
                                                font=10, anchor="n")
        standard_calibration_date_3.grid(row=3, column=4)
        standard_calibration_date_3.config(width=date_of_calibration_width)
        standard_calibration_due_date_3 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_3,
                                                    font=10, anchor="n")
        standard_calibration_due_date_3.grid(row=3, column=5)
        standard_calibration_due_date_3.config(width=calibration_due_date_width)

        # Asset Number 4
        lbl_asset_number_4 = ttk.Label(calibration_standard_frame, text="4) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_4.grid(row=4, column=0, pady=5, padx=5)
        calibration_standard_4 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_4_variable,
                                           font=('arial', 12))
        calibration_standard_4.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_4_variable))
        calibration_standard_4.grid(row=4, column=1)
        calibration_standard_4.config(width=15)

        # Asset Number 4 - Description, Serial, Cal, Due Date
        standard_description_4 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_4, font=10,
                                           anchor="n")
        standard_description_4.grid(row=4, column=2)
        standard_description_4.config(width=description_width)
        standard_serial_number_4 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_4, font=10,
                                             anchor="n")
        standard_serial_number_4.grid(row=4, column=3)
        standard_serial_number_4.config(width=serial_number_width)
        standard_calibration_date_4 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_4,
                                                font=10, anchor="n")
        standard_calibration_date_4.grid(row=4, column=4)
        standard_calibration_date_4.config(width=date_of_calibration_width)
        standard_calibration_due_date_4 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_4,
                                                    font=10, anchor="n")
        standard_calibration_due_date_4.grid(row=4, column=5)
        standard_calibration_due_date_4.config(width=calibration_due_date_width)

        # Asset Number 5
        lbl_asset_number_5 = ttk.Label(calibration_standard_frame, text="5) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_5.grid(row=5, column=0, pady=5, padx=5)
        calibration_standard_5 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_5_variable,
                                           font=('arial', 12))
        calibration_standard_5.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_5_variable))
        calibration_standard_5.grid(row=5, column=1)
        calibration_standard_5.config(width=15)

        # Asset Number 5 - Description, Serial, Cal, Due Date
        standard_description_5 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_5, font=10,
                                           anchor="n")
        standard_description_5.grid(row=5, column=2)
        standard_description_5.config(width=description_width)
        standard_serial_number_5 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_5, font=10,
                                             anchor="n")
        standard_serial_number_5.grid(row=5, column=3)
        standard_serial_number_5.config(width=serial_number_width)
        standard_calibration_date_5 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_5,
                                                font=10, anchor="n")
        standard_calibration_date_5.grid(row=5, column=4)
        standard_calibration_date_5.config(width=date_of_calibration_width)
        standard_calibration_due_date_5 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_5,
                                                    font=10, anchor="n")
        standard_calibration_due_date_5.grid(row=5, column=5)
        standard_calibration_due_date_5.config(width=calibration_due_date_width)

        # Asset Number 6
        lbl_asset_number_6 = ttk.Label(calibration_standard_frame, text="6) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_6.grid(row=6, column=0, pady=5, padx=5)
        calibration_standard_6 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_6_variable,
                                           font=('arial', 12))
        calibration_standard_6.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_6_variable))
        calibration_standard_6.grid(row=6, column=1)
        calibration_standard_6.config(width=15)

        # Asset Number 6 - Description, Serial, Cal, Due Date
        standard_description_6 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_6, font=10,
                                           anchor="n")
        standard_description_6.grid(row=6, column=2)
        standard_description_6.config(width=description_width)
        standard_serial_number_6 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_6, font=10,
                                             anchor="n")
        standard_serial_number_6.grid(row=6, column=3)
        standard_serial_number_6.config(width=serial_number_width)
        standard_calibration_date_6 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_6,
                                                font=10, anchor="n")
        standard_calibration_date_6.grid(row=6, column=4)
        standard_calibration_date_6.config(width=date_of_calibration_width)
        standard_calibration_due_date_6 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_6,
                                                    font=10, anchor="n")
        standard_calibration_due_date_6.grid(row=6, column=5)
        standard_calibration_due_date_6.config(width=calibration_due_date_width)

        # Asset Number 7
        lbl_asset_number_7 = ttk.Label(calibration_standard_frame, text="7) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_7.grid(row=7, column=0, pady=5, padx=5)
        calibration_standard_7 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_7_variable,
                                           font=('arial', 12))
        calibration_standard_7.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_7_variable))
        calibration_standard_7.grid(row=7, column=1)
        calibration_standard_7.config(width=15)

        # Asset Number 7 - Description, Serial, Cal, Due Date
        standard_description_7 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_7, font=10,
                                           anchor="n")
        standard_description_7.grid(row=7, column=2)
        standard_description_7.config(width=description_width)
        standard_serial_number_7 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_7, font=10,
                                             anchor="n")
        standard_serial_number_7.grid(row=7, column=3)
        standard_serial_number_7.config(width=serial_number_width)
        standard_calibration_date_7 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_7,
                                                font=10, anchor="n")
        standard_calibration_date_7.grid(row=7, column=4)
        standard_calibration_date_7.config(width=date_of_calibration_width)
        standard_calibration_due_date_7 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_7,
                                                    font=10, anchor="n")
        standard_calibration_due_date_7.grid(row=7, column=5)
        standard_calibration_due_date_7.config(width=calibration_due_date_width)

        # Asset Number 8
        lbl_asset_number_8 = ttk.Label(calibration_standard_frame, text="8) Asset Number:", font=('arial', 12),
                                       anchor="n")
        lbl_asset_number_8.grid(row=8, column=0, pady=5, padx=5)
        calibration_standard_8 = ttk.Entry(calibration_standard_frame, textvariable=calibration_standard_8_variable,
                                           font=('arial', 12))
        calibration_standard_8.bind("<KeyRelease>", lambda event: acc.all_caps(calibration_standard_8_variable))
        calibration_standard_8.grid(row=8, column=1)
        calibration_standard_8.config(width=15)

        # Asset Number 8 - Description, Serial, Cal, Due Date
        standard_description_8 = ttk.Label(calibration_standard_frame,
                                           text=LIMSVarConfig.calibration_standard_equipment_description_8, font=10,
                                           anchor="n")
        standard_description_8.grid(row=8, column=2)
        standard_description_8.config(width=description_width)
        standard_serial_number_8 = ttk.Label(calibration_standard_frame,
                                             text=LIMSVarConfig.calibration_standard_serial_number_8, font=10,
                                             anchor="n")
        standard_serial_number_8.grid(row=8, column=3)
        standard_serial_number_8.config(width=serial_number_width)
        standard_calibration_date_8 = ttk.Label(calibration_standard_frame,
                                                text=LIMSVarConfig.calibration_standard_equipment_calibration_date_8,
                                                font=10, anchor="n")
        standard_calibration_date_8.grid(row=8, column=4)
        standard_calibration_date_8.config(width=date_of_calibration_width)
        standard_calibration_due_date_8 = ttk.Label(calibration_standard_frame,
                                                    text=LIMSVarConfig.calibration_standard_equipment_due_date_8,
                                                    font=10, anchor="n")
        standard_calibration_due_date_8.grid(row=8, column=5)
        standard_calibration_due_date_8.config(width=calibration_due_date_width)

        # These label is a dummy label that does nothing but help format the \
        # frame
        dummy = ttk.Label(calibration_standard_frame)
        dummy.grid(row=1, column=6)
        dummy.config(width=2)

        # .....................Button Labels and Buttons.........................#

        btn_load_calibration_standard_information = ttk.Button(calibration_standard_frame, text="Load", width=20,
                                                               command=lambda: self.load_calibration_standards())
        btn_load_calibration_standard_information.bind("<Return>", lambda event: self.load_calibration_standards())
        btn_load_calibration_standard_information.grid(pady=5, row=9, column=0, columnspan=7)

        btn_back_out = ttk.Button(CalibrationStandardDetails, text="Back", width=20,
                                  command=lambda: self.device_under_test_calibration_specifications(CalibrationStandardDetails))
        btn_back_out.bind("<Return>",
                          lambda event: self.device_under_test_calibration_specifications(CalibrationStandardDetails))
        btn_back_out.grid(pady=5, row=7, column=1, columnspan=2)

        btn_import_calibration_data = ttk.Button(CalibrationStandardDetails, text="Import Data", width=20,
                                                 command=lambda: self.calibration_standards_equipment_import_input_verification())
        btn_import_calibration_data.bind("<Return>",
                                         lambda event: self.calibration_standards_equipment_import_input_verification())
        btn_import_calibration_data.grid(pady=5, row=7, column=2, columnspan=2)

        btn_perform_calibration = ttk.Button(CalibrationStandardDetails, text="Perform Calibration", width=20,
                                             command=lambda: self.calibration_standards_equipment_input_verification())
        btn_perform_calibration.bind("<Return>",
                                     lambda event: self.calibration_standards_equipment_input_verification())
        btn_perform_calibration.grid(pady=5, row=7, column=3, columnspan=2)

    # -----------------------------------------------------------------------#

    # Command to Load Calibration Standards from Database
    def load_calibration_standards(self):
        global todays_date

        if calibration_standard_1.get() == "":
            tm.showerror("Calibration Standard Import Error", "No calibration standard asset number provided in \
entry box. Please provide a valid asset number for the calibration equipment used in the calibration.")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_1.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date = LIMSVarConfig.cal_equip_due_date[i]
            standard_description.config(text=LIMSVarConfig.calibration_standard_equipment_description)
            standard_serial_number.config(text=LIMSVarConfig.calibration_standard_serial_number)
            standard_calibration_date.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date)

            now = datetime.datetime.now()
            todays_date = now.strftime("%m-%d-%Y").replace("-", "/")
            due_date_1 = LIMSVarConfig.calibration_standard_equipment_due_date

            if datetime.datetime.strptime(due_date_1, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date.config(text=LIMSVarConfig.calibration_standard_equipment_due_date)
                standard_calibration_due_date.config(background="systemMenu")
            else:
                standard_calibration_due_date.config(text=LIMSVarConfig.calibration_standard_equipment_due_date)
                standard_calibration_due_date.config(background="red")

        if calibration_standard_2.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_2 = ""
            LIMSVarConfig.calibration_standard_serial_number_2 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_2 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_2 = ""
            standard_description_2.config(text=LIMSVarConfig.calibration_standard_equipment_description_2)
            standard_serial_number_2.config(text=LIMSVarConfig.calibration_standard_serial_number_2)
            standard_calibration_date_2.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_2)
            standard_calibration_due_date_2.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_2)
            standard_calibration_due_date_2.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_2.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_2 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_2 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_2 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_2 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_2.config(text=LIMSVarConfig.calibration_standard_equipment_description_2)
            standard_serial_number_2.config(text=LIMSVarConfig.calibration_standard_serial_number_2)
            standard_calibration_date_2.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_2)

            due_date_2 = LIMSVarConfig.calibration_standard_equipment_due_date_2

            if datetime.datetime.strptime(due_date_2, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_2.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_2)
                standard_calibration_due_date_2.config(background="systemMenu")
            else:
                standard_calibration_due_date_2.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_2)
                standard_calibration_due_date_2.config(background="red")

        if calibration_standard_3.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_3 = ""
            LIMSVarConfig.calibration_standard_serial_number_3 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_3 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_3 = ""
            standard_description_3.config(text=LIMSVarConfig.calibration_standard_equipment_description_3)
            standard_serial_number_3.config(text=LIMSVarConfig.calibration_standard_serial_number_3)
            standard_calibration_date_3.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_3)
            standard_calibration_due_date_3.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_3)
            standard_calibration_due_date_3.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_3.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_3 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_3 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_3 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_3 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_3.config(text=LIMSVarConfig.calibration_standard_equipment_description_3)
            standard_serial_number_3.config(text=LIMSVarConfig.calibration_standard_serial_number_3)
            standard_calibration_date_3.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_3)

            due_date_3 = LIMSVarConfig.calibration_standard_equipment_due_date_3

            if datetime.datetime.strptime(due_date_3, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_3.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_3)
                standard_calibration_due_date_3.config(background="systemMenu")
            else:
                standard_calibration_due_date_3.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_3)
                standard_calibration_due_date_3.config(background="red")

        if calibration_standard_4.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_4 = ""
            LIMSVarConfig.calibration_standard_serial_number_4 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_4 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_4 = ""
            standard_description_4.config(text=LIMSVarConfig.calibration_standard_equipment_description_4)
            standard_serial_number_4.config(text=LIMSVarConfig.calibration_standard_serial_number_4)
            standard_calibration_date_4.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_4)
            standard_calibration_due_date_4.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_4)
            standard_calibration_due_date_4.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_4.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_4 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_4 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_4 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_4 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_4.config(text=LIMSVarConfig.calibration_standard_equipment_description_4)
            standard_serial_number_4.config(text=LIMSVarConfig.calibration_standard_serial_number_4)
            standard_calibration_date_4.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_4)

            due_date_4 = LIMSVarConfig.calibration_standard_equipment_due_date_4

            if datetime.datetime.strptime(due_date_4, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_4.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_4)
                standard_calibration_due_date_4.config(background="systemMenu")
            else:
                standard_calibration_due_date_4.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_4)
                standard_calibration_due_date_4.config(background="red")

        if calibration_standard_5.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_5 = ""
            LIMSVarConfig.calibration_standard_serial_number_5 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_5 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_5 = ""
            standard_description_5.config(text=LIMSVarConfig.calibration_standard_equipment_description_5)
            standard_serial_number_5.config(text=LIMSVarConfig.calibration_standard_serial_number_5)
            standard_calibration_date_5.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_5)
            standard_calibration_due_date_5.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_5)
            standard_calibration_due_date_5.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_5.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_5 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_5 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_5 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_5 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_5.config(text=LIMSVarConfig.calibration_standard_equipment_description_5)
            standard_serial_number_5.config(text=LIMSVarConfig.calibration_standard_serial_number_5)
            standard_calibration_date_5.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_5)

            due_date_5 = LIMSVarConfig.calibration_standard_equipment_due_date_5

            if datetime.datetime.strptime(due_date_5, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_5.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_5)
                standard_calibration_due_date_5.config(background="systemMenu")
            else:
                standard_calibration_due_date_5.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_5)
                standard_calibration_due_date_5.config(background="red")

        if calibration_standard_6.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_6 = ""
            LIMSVarConfig.calibration_standard_serial_number_6 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_6 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_6 = ""
            standard_description_6.config(text=LIMSVarConfig.calibration_standard_equipment_description_6)
            standard_serial_number_6.config(text=LIMSVarConfig.calibration_standard_serial_number_6)
            standard_calibration_date_6.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_6)
            standard_calibration_due_date_6.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_6)
            standard_calibration_due_date_6.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_6.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_6 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_6 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_6 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_6 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_6.config(text=LIMSVarConfig.calibration_standard_equipment_description_6)
            standard_serial_number_6.config(text=LIMSVarConfig.calibration_standard_serial_number_6)
            standard_calibration_date_6.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_6)

            due_date_6 = LIMSVarConfig.calibration_standard_equipment_due_date_6

            if datetime.datetime.strptime(due_date_6, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_6.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_6)
                standard_calibration_due_date_6.config(background="systemMenu")
            else:
                standard_calibration_due_date_6.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_6)
                standard_calibration_due_date_6.config(background="red")

        if calibration_standard_7.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_7 = ""
            LIMSVarConfig.calibration_standard_serial_number_7 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_7 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_7 = ""
            standard_description_7.config(text=LIMSVarConfig.calibration_standard_equipment_description_7)
            standard_serial_number_7.config(text=LIMSVarConfig.calibration_standard_serial_number_7)
            standard_calibration_date_7.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_7)
            standard_calibration_due_date_7.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_7)
            standard_calibration_due_date_7.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_7.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_7 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_7 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_7 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_7 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_7.config(text=LIMSVarConfig.calibration_standard_equipment_description_7)
            standard_serial_number_7.config(text=LIMSVarConfig.calibration_standard_serial_number_7)
            standard_calibration_date_7.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_7)

            due_date_7 = LIMSVarConfig.calibration_standard_equipment_due_date_7

            if datetime.datetime.strptime(due_date_7, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_7.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_7)
                standard_calibration_due_date_7.config(background="systemMenu")
            else:
                standard_calibration_due_date_7.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_7)
                standard_calibration_due_date_7.config(background="red")

        if calibration_standard_8.get() == "":
            LIMSVarConfig.calibration_standard_equipment_description_8 = ""
            LIMSVarConfig.calibration_standard_serial_number_8 = ""
            LIMSVarConfig.calibration_standard_equipment_calibration_date_8 = ""
            LIMSVarConfig.calibration_standard_equipment_due_date_8 = ""
            standard_description_8.config(text=LIMSVarConfig.calibration_standard_equipment_description_8)
            standard_serial_number_8.config(text=LIMSVarConfig.calibration_standard_serial_number_8)
            standard_calibration_date_8.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_8)
            standard_calibration_due_date_8.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_8)
            standard_calibration_due_date_8.config(background="systemMenu")
        else:
            for i in range(0, len(LIMSVarConfig.cal_equip_item_no)):
                if calibration_standard_8.get() != LIMSVarConfig.cal_equip_asset_no[i]:
                    i += 1
                else:
                    LIMSVarConfig.calibration_standard_equipment_description_8 = LIMSVarConfig.cal_equip_descrip[i]
                    LIMSVarConfig.calibration_standard_serial_number_8 = LIMSVarConfig.cal_equip_serial_no[i]
                    LIMSVarConfig.calibration_standard_equipment_calibration_date_8 = LIMSVarConfig.cal_equip_cal_date[i]
                    LIMSVarConfig.calibration_standard_equipment_due_date_8 = LIMSVarConfig.cal_equip_due_date[i]
            standard_description_8.config(text=LIMSVarConfig.calibration_standard_equipment_description_8)
            standard_serial_number_8.config(text=LIMSVarConfig.calibration_standard_serial_number_8)
            standard_calibration_date_8.config(text=LIMSVarConfig.calibration_standard_equipment_calibration_date_8)

            due_date_8 = LIMSVarConfig.calibration_standard_equipment_due_date_8

            if datetime.datetime.strptime(due_date_8, '%m/%d/%Y') > datetime.datetime.strptime(todays_date, '%m/%d/%Y'):
                standard_calibration_due_date_8.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_8)
                standard_calibration_due_date_8.config(background="systemMenu")
            else:
                standard_calibration_due_date_8.config(text=LIMSVarConfig.calibration_standard_equipment_due_date_8)
                standard_calibration_due_date_8.config(background="red")

    # -----------------------------------------------------------------------#

    # Command to ensure that fields have been filled prior to advancing to next window
    def calibration_standards_equipment_input_verification(self):

        if calibration_standard_1.get() == "" and calibration_standard_2.get() == "" and \
                calibration_standard_3.get() == "" and calibration_standard_4.get() == "" and \
                calibration_standard_5.get() == "":
            tm.showerror("Error", "Please provide at least two calibration standards used during calibration.")
        elif calibration_standard_1.get() == "":
            tm.showerror("Error", 'Please provide the first calibration standard asset number used in the first \
asset number entry field provided.')
        elif calibration_standard_2.get() == "":
            tm.showerror("Error", "Please provide at least two calibration standards used during calibration.")
        else:
            self.perform_calibration_on_device_under_test(CalibrationStandardDetails)

    # -----------------------------------------------------------------------#

    # Command to ensure that fields have been filled prior to advancing to next window
    def calibration_standards_equipment_import_input_verification(self):
        from LIMSDataImport import AppDataImportModule
        adim = AppDataImportModule()

        self.__init__()

        if calibration_standard_1.get() == "" and calibration_standard_2.get() == "" and \
                calibration_standard_3.get() == "" and calibration_standard_4.get() == "" and \
                calibration_standard_5.get() == "":
            tm.showerror("Error", "Please provide at least two calibration standards used during calibration.")
        elif calibration_standard_1.get() == "":
            tm.showerror("Error", 'Please provide the first calibration standard asset number used in the first \
asset number entry field provided.')
        elif calibration_standard_2.get() == "":
            tm.showerror("Error", "Please provide at least two calibration standards used during calibration.")
        else:
            adim.data_import_selection(CalibrationStandardDetails)

    # -----------------------------------------------------------------------#

    # Command to Send User to Perform Calibration Section of Cal Certificate
    def perform_calibration_on_device_under_test(self, window):
        window.withdraw()

        global DUTCalibrationWindow, device_under_test_full_scale_calculated, \
            second_device_under_test_full_scale_calculated, device_under_test_full_scale_range_calculated, \
            second_device_under_test_full_scale_range, data_table, second_data_table, device_under_test_reading_entry, \
            device_reading_list, reference_reading_entry, reference_reading_list, nominal_value_entry, \
            nominal_input_list, second_dut_reading_list, second_reference_reading_list, \
            second_device_under_test_reading_entry, second_reference_reading_entry, btn_complete_calibration_process, second_w, w, second_h, height, width

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        aph = AppHelpWindows()

        DUTCalibrationWindow = Toplevel()
        DUTCalibrationWindow.title("Calibration Data Input")
        DUTCalibrationWindow.iconbitmap("Required Images\\DwyerLogo.ico")

        if dut_output_type.get() == "Single":
            height = 720
            width = 630
        elif dut_output_type.get() == "Transmitter":
            height = 720
            width = 760
        elif dut_output_type.get() == "Dual":
            height = 720
            width = 1190

        screen_width = DUTCalibrationWindow.winfo_screenwidth()
        screen_height = DUTCalibrationWindow.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        DUTCalibrationWindow.geometry("%dx%d+%d+%d" % (width, height, x, y))
        DUTCalibrationWindow.focus_force()
        DUTCalibrationWindow.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(DUTCalibrationWindow))

        # ......................Full Scale Calculation............................#

        if device_under_test_output_type is not None:
            if device_under_test_minimum < 0 or device_under_test_minimum == 0:
                device_under_test_full_scale_calculated = (abs(float(device_under_test_minimum)) +
                                                           float(device_under_test_maximum))
            else:
                device_under_test_full_scale_calculated = float(device_under_test_maximum)
                device_under_test_full_scale_range_calculated = (float(device_under_test_maximum) -
                                                                 float(device_under_test_minimum))

        if device_under_test_output_type == "Dual" or device_under_test_output_type == "Transmitter":
            if second_device_under_test_minimum <= 0:
                second_device_under_test_full_scale_calculated = (abs(float(second_device_under_test_minimum)) +
                                                                  float(second_device_under_test_maximum))
            else:
                second_device_under_test_full_scale_calculated = float(second_device_under_test_maximum)
                second_device_under_test_full_scale_range = (float(second_device_under_test_maximum) -
                                                             float(second_device_under_test_minimum))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(DUTCalibrationWindow, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(DUTCalibrationWindow)),
                                           ("Calibration Standard Information",
                                            lambda: self.calibration_standards_selection(DUTCalibrationWindow)),
                                           ("Logout", lambda: acc.software_signout(DUTCalibrationWindow)),
                                           ("Quit", lambda: acc.software_close(DUTCalibrationWindow))])
        menubar.add_menu("Help", commands=[("Help", lambda: aph.calibration_data_input_help())])

        # .......................Data Table Creation Criteria.....................#

        h = int(dut_number_of_test_points) + int(1)

        if device_under_test_output_type == "Dual":
            second_h = int(second_device_under_test_number_of_test_points) + int(1)
            second_w = 14

        if device_under_test_output_type != "Transmitter":
            w = 7
        elif device_under_test_output_type == "Transmitter":
            w = 8

        # for i in range(h):  # Rows
        #     for j in range(w):  # Columns
        #         data_table = ttk.Label(DUTCalibrationWindow, text="", anchor="n")
        #         data_table.grid(row=i, column=j)
        for k in range(h):
            test_step_number = ttk.Label(DUTCalibrationWindow, text=k, anchor="n")
            test_step_number.grid(row=k, column=0)

        device_reading_list = []
        reference_reading_list = []

        if device_under_test_output_type == "Transmitter":
            nominal_input_list = []

        if device_under_test_output_type != "Transmitter":
            for l in range(h - 1):
                device_under_test_reading_entry = ttk.Entry(DUTCalibrationWindow)
                device_under_test_reading_entry.config(width=measurement_discipline_type_width)
                device_under_test_reading_entry.grid(row=l + 1, column=1)
                device_reading_list.append(device_under_test_reading_entry)
                reference_reading_entry = ttk.Entry(DUTCalibrationWindow)
                reference_reading_entry.config(width=measurement_discipline_type_width)
                reference_reading_entry.grid(row=l + 1, column=3)
                reference_reading_list.append(reference_reading_entry)
                data_table = ttk.Label(DUTCalibrationWindow, text=units_of_measure, anchor="n")
                data_table.grid(row=l + 1, column=2)
                data_table = ttk.Label(DUTCalibrationWindow, text=units_of_measure, anchor="n")
                data_table.grid(row=l + 1, column=4)
            data_table = ttk.Label(DUTCalibrationWindow, text="Test Step Number", anchor="n")
            data_table.grid(row=0, column=0, columnspan=1, padx=10)
            data_table = ttk.Label(DUTCalibrationWindow, text="DUT Reading", anchor="n")
            data_table.grid(row=0, column=1, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
            data_table.grid(row=0, column=2, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Reference Reading", anchor="n")
            data_table.grid(row=0, column=3, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
            data_table.grid(row=0, column=4, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Calculated Error", anchor="n")
            data_table.grid(row=0, column=5, columnspan=1)
            pass_fail_criteria = ttk.Label(DUTCalibrationWindow, text="Pass/Fail", anchor="n")
            pass_fail_criteria.config(width=measurement_discipline_type_width)
            pass_fail_criteria.grid(row=0, column=6, columnspan=1)

        elif device_under_test_output_type == "Transmitter":
            for l in range(h - 1):
                nominal_value_entry = ttk.Entry(DUTCalibrationWindow)
                nominal_value_entry.config(width=measurement_discipline_type_width)
                nominal_value_entry.grid(row=l + 1, column=1)
                nominal_input_list.append(nominal_value_entry)
                reference_reading_entry = ttk.Entry(DUTCalibrationWindow)
                reference_reading_entry.config(width=measurement_discipline_type_width)
                reference_reading_entry.grid(row=l + 1, column=3)
                reference_reading_list.append(reference_reading_entry)
                device_under_test_reading_entry = ttk.Entry(DUTCalibrationWindow)
                device_under_test_reading_entry.config(width=measurement_discipline_type_width)
                device_under_test_reading_entry.grid(row=l + 1, column=5)
                device_reading_list.append(device_under_test_reading_entry)
                data_table = ttk.Label(DUTCalibrationWindow, text=units_of_measure, anchor="n")
                data_table.grid(row=l + 1, column=2)
                data_table = ttk.Label(DUTCalibrationWindow, text=units_of_measure, anchor="n")
                data_table.grid(row=l + 1, column=4)
                data_table = ttk.Label(DUTCalibrationWindow, text=second_units_of_measure, anchor="n")
                data_table.grid(row=l + 1, column=6)
            data_table = ttk.Label(DUTCalibrationWindow, text="Test Step Number", anchor="n")
            data_table.grid(row=0, column=0, columnspan=1, padx=10)
            data_table = ttk.Label(DUTCalibrationWindow, text="Nominal Value", anchor="n")
            data_table.grid(row=0, column=1, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
            data_table.grid(row=0, column=2, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Reference Reading", anchor="n")
            data_table.grid(row=0, column=3, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
            data_table.grid(row=0, column=4, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="DUT Output", anchor="n")
            data_table.grid(row=0, column=5, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
            data_table.grid(row=0, column=6, columnspan=1)
            data_table = ttk.Label(DUTCalibrationWindow, text="Calculated Error", anchor="n")
            data_table.grid(row=0, column=7, columnspan=1)
            pass_fail_criteria = ttk.Label(DUTCalibrationWindow, text="Pass/Fail", anchor="n")
            pass_fail_criteria.config(width=measurement_discipline_type_width)
            pass_fail_criteria.grid(row=0, column=8, columnspan=1)

        if device_under_test_output_type == "Dual":
            for i in range(second_h):
                for j in range(w + 1, second_w):
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="", anchor="n")
                    second_data_table.grid(row=i, column=j)
                    for k in range(second_h):
                        second_data_table = ttk.Label(DUTCalibrationWindow, text=k, anchor="n")
                        second_data_table.grid(row=k, column=7)
                    second_dut_reading_list = []
                    second_reference_reading_list = []
                    for l in range(second_h - 1):
                        second_device_under_test_reading_entry = ttk.Entry(DUTCalibrationWindow)
                        second_device_under_test_reading_entry.config(width=measurement_discipline_type_width)
                        second_device_under_test_reading_entry.grid(row=l + 1, column=8)
                        second_dut_reading_list.append(second_device_under_test_reading_entry)
                        second_reference_reading_entry = ttk.Entry(DUTCalibrationWindow)
                        second_reference_reading_entry.config(width=measurement_discipline_type_width)
                        second_reference_reading_entry.grid(row=l + 1, column=10)
                        second_reference_reading_list.append(second_reference_reading_entry)
                        second_data_table = ttk.Label(DUTCalibrationWindow, text=second_units_of_measure, anchor="n")
                        second_data_table.grid(row=l + 1, column=9)
                        second_data_table = ttk.Label(DUTCalibrationWindow, text=second_units_of_measure, anchor="n")
                        second_data_table.grid(row=l + 1, column=11)
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="Test Step Number", anchor="n")
                    second_data_table.grid(row=0, column=7, columnspan=1, padx=10)
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="DUT Reading", anchor="n")
                    second_data_table.grid(row=0, column=8, columnspan=1)
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
                    second_data_table.grid(row=0, column=9, columnspan=1)
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="Reference Reading", anchor="n")
                    second_data_table.grid(row=0, column=10, columnspan=1)
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="Units", anchor="n")
                    second_data_table.grid(row=0, column=11, columnspan=1)
                    second_data_table = ttk.Label(DUTCalibrationWindow, text="Calculated Error", anchor="n")
                    second_data_table.grid(row=0, column=12, columnspan=1)
                    second_pass_fail_criteria = ttk.Label(DUTCalibrationWindow, text="Pass/Fail", anchor="n")
                    second_pass_fail_criteria.config(width=measurement_discipline_type_width)
                    second_pass_fail_criteria.grid(row=0, column=13, columnspan=1)

        # ...............................Buttons..................................#

        btn_back_out_calibration = ttk.Button(DUTCalibrationWindow, text="Back",
                                              width=measurement_discipline_type_width,
                                              command=lambda: self.calibration_standards_selection(DUTCalibrationWindow)
                                              )
        btn_back_out_calibration.bind("<Return>",
                                      lambda event: self.calibration_standards_selection(DUTCalibrationWindow))
        if device_under_test_output_type != "Dual":
            btn_back_out_calibration.grid(pady=5, row=h+1, column=1)
        elif device_under_test_output_type == "Dual":
            btn_back_out_calibration.grid(pady=5, row=h+1, column=3)

        btn_calculate_error = ttk.Button(DUTCalibrationWindow, text="Calculate",
                                         width=measurement_discipline_type_width,
                                         command=lambda: self.calculate_calibration_measurement_error())
        btn_calculate_error.bind("<Return>", lambda event: self.calculate_calibration_measurement_error())
        if device_under_test_output_type != "Dual":
            btn_calculate_error.grid(pady=5, row=h+1, column=3)
        elif device_under_test_output_type == "Dual":
            btn_calculate_error.grid(pady=5, row=h+1, column=6, columnspan=2)

        btn_complete_calibration_process = ttk.Button(DUTCalibrationWindow, text="Complete",
                                                      width=measurement_discipline_type_width,
                                                      command=lambda: self.complete_calibration_process_step_one())
        btn_complete_calibration_process.bind("<Return>", lambda event: self.complete_calibration_process_step_one())
        if device_under_test_output_type != "Dual":
            btn_complete_calibration_process.grid(pady=5, row=h+1, column=5)
            btn_complete_calibration_process.config(state=DISABLED)
        elif device_under_test_output_type == "Dual":
            btn_complete_calibration_process.grid(pady=5, row=h+1, column=10)
            btn_complete_calibration_process.config(state=DISABLED)

    # -----------------------------------------------------------------------#

    # Command to Update Calculated Error Field and Pass/Fail Criteria for
    # DUT Calibration Window/Data Acquisition Phase of Calibration

    def calculate_calibration_measurement_error(self):

        global total_error_band_list, second_total_error_band_list, pass_fail_condition_list, \
            dual_pass_fail_condition_list, transmitter_pass_fail_condition_list, device_under_test_reading_entry, \
            reference_reading_entry, nominal_value_entry, second_error_4, second_error_3, second_error_2, second_error_1, second_error_band_4_list, second_error_band_3_list, second_error_band_2_list_, second_error_band_list, second_calculated_difference_list, second_x_axis_scaling_list, judgement_condition, error_4_value, error_3_value, error_2_value, error_1_value, target_supply_span, judgement_condition, error_4_value, error_3_value, error_2_value, error_1_value

        # ......................Array Initialization and Value Definitions.....................#

        btn_complete_calibration_process.config(state=NORMAL)

        # Color Configurations for Pass/Fail Conditions
        s = ttk.Style()
        s.configure('pass.TLabel', width=measurement_discipline_type_width, background="green")

        s = ttk.Style()
        s.configure('fail.TLabel', width=measurement_discipline_type_width, background="red")

        pass_fail_condition_list = []
        dual_pass_fail_condition_list = []
        transmitter_pass_fail_condition_list = []

        i = 0
        x_axis_scaling_list = []
        calculated_difference_list = []
        error_band_list = []
        error_band_2_list = []
        error_band_3_list = []
        error_band_4_list = []
        total_error_band_list = []

        if device_under_test_output_type == "Dual":
            second_x_axis_scaling_list = []
            second_calculated_difference_list = []
            second_error_band_list = []
            second_error_band_2_list_ = []
            second_error_band_3_list = []
            second_error_band_4_list = []
            second_total_error_band_list = []

        if device_under_test_output_type == "Transmitter":
            if device_under_test_minimum <= 0:
                target_supply_span = abs(float(device_under_test_minimum)) + float(device_under_test_maximum)
            else:
                target_supply_span = float(device_under_test_maximum) - float(device_under_test_minimum)

        # .........................Generate Graph for Calibration Data........................#

        # /////////////////////////////Single and Dual Output Data////////////////////////////#

        if device_under_test_output_type != "Transmitter":
            for device_under_test_reading_entry, reference_reading_entry in zip(device_reading_list,
                                                                                reference_reading_list):
                x_axis_scaling_list.append(device_under_test_reading_entry.get())
                calculated_measured_difference = (float(device_under_test_reading_entry.get()) -
                                                  float(reference_reading_entry.get()))
                calculated_measured_difference_value = round(calculated_measured_difference, int(reference_resolution))
                calculated_difference_list.append(calculated_measured_difference_value)

                if device_under_test_specification_type == "Actual Error":
                    error_band_list.append(device_under_test_specification_value)
                    error_1_value = device_under_test_specification_value

                elif device_under_test_specification_type == "% of Full Scale":
                    error_band_list.append(((float(device_under_test_specification_value)) / 100) *
                                           device_under_test_full_scale_calculated)
                    error_1_value = (((float(device_under_test_specification_value)) / 100) *
                                     device_under_test_full_scale_calculated)

                elif device_under_test_specification_type == "% of FS (Range)":
                    error_band_list.append(((float(device_under_test_specification_value)) / 100) *
                                           device_under_test_full_scale_range_calculated)
                    error_1_value = (((float(device_under_test_specification_value)) / 100) *
                                     device_under_test_full_scale_range_calculated)

                elif device_under_test_specification_type == "% of Reading":
                    error_band_list.append(((float(device_under_test_specification_value)) / 100) *
                                           float(reference_reading_entry.get()))
                    error_1_value = (((float(device_under_test_specification_value)) / 100) *
                                     float(reference_reading_entry.get()))

                if device_under_test_specification_value_1 is None or device_under_test_specification_value_1 == "" or \
                        device_under_test_specification_value_1 == " ":
                    pass
                else:
                    if device_under_test_specification_type_1 == "Actual Error":
                        error_band_2_list.append(device_under_test_specification_value_1)
                        error_2_value = device_under_test_specification_value_1

                    elif device_under_test_specification_type_1 == "% of Full Scale":
                        error_band_2_list.append(((float(device_under_test_specification_value_1)) / 100) *
                                                 device_under_test_full_scale_calculated)
                        error_2_value = (((float(device_under_test_specification_value_1)) / 100) *
                                         device_under_test_full_scale_calculated)

                    elif device_under_test_specification_type_1 == "% of FS (Range)":
                        error_band_2_list.append(((float(device_under_test_specification_value_1)) / 100) *
                                                 device_under_test_full_scale_range_calculated)
                        error_2_value = (((float(device_under_test_specification_value_1)) / 100) *
                                         device_under_test_full_scale_range_calculated)

                    elif device_under_test_specification_type_1 == "% of Reading":
                        error_band_2_list.append(((float(device_under_test_specification_value_1)) / 100) *
                                                 float(reference_reading_entry.get()))
                        error_2_value = (((float(device_under_test_specification_value_1)) / 100) *
                                         float(reference_reading_entry.get()))

                if device_under_test_specification_value_2 is None or device_under_test_specification_value_2 == "" or \
                        device_under_test_specification_value_2 == " ":
                    pass
                else:
                    if device_under_test_specification_type_2 == "Actual Error":
                        error_band_3_list.append(device_under_test_specification_value_2)
                        error_3_value = device_under_test_specification_value_2

                    elif device_under_test_specification_type_2 == "% of Full Scale":
                        error_band_3_list.append(((float(device_under_test_specification_value_2)) / 100) *
                                                 device_under_test_full_scale_calculated)
                        error_3_value = (((float(device_under_test_specification_value_2)) / 100) *
                                         device_under_test_full_scale_calculated)

                    elif device_under_test_specification_type_2 == "% of FS (Range)":
                        error_band_3_list.append(((float(device_under_test_specification_value_2)) / 100) *
                                                 device_under_test_full_scale_range_calculated)
                        error_3_value = (((float(device_under_test_specification_value_2)) / 100) *
                                         device_under_test_full_scale_range_calculated)

                    elif device_under_test_specification_type_2 == "% of Reading":
                        error_band_3_list.append(((float(device_under_test_specification_value_2)) / 100) *
                                                 float(reference_reading_entry.get()))
                        error_3_value = (((float(device_under_test_specification_value_2)) / 100) *
                                         float(reference_reading_entry.get()))

                if device_under_test_specification_value_3 is None or device_under_test_specification_value_3 == "" or \
                        device_under_test_specification_value_3 == " ":
                    pass
                else:
                    if device_under_test_specification_type_3 == "Actual Error":
                        error_band_4_list.append(device_under_test_specification_value_3)
                        error_4_value = device_under_test_specification_value_3

                    elif device_under_test_specification_type_3 == "% of Full Scale":
                        error_band_4_list.append(((float(device_under_test_specification_value_3)) / 100) *
                                                 device_under_test_full_scale_calculated)
                        error_4_value = (((float(device_under_test_specification_value_3)) / 100) *
                                         device_under_test_full_scale_calculated)

                    elif device_under_test_specification_type_3 == "% of FS (Range)":
                        error_band_4_list.append(((float(device_under_test_specification_value_3)) / 100) *
                                                 device_under_test_full_scale_range_calculated)
                        error_4_value = (((float(device_under_test_specification_value_3)) / 100) *
                                         device_under_test_full_scale_range_calculated)

                    elif device_under_test_specification_type_3 == "% of Reading":
                        error_band_4_list.append(((float(device_under_test_specification_value_3)) / 100) *
                                                 float(reference_reading_entry.get()))
                        error_4_value = (((float(device_under_test_specification_value_3)) / 100) *
                                         float(reference_reading_entry.get()))

                # ///////////////////Single and Dual Output Error Band Determination//////////////////////#

                if device_under_test_specification_operator_1 == "" and \
                        device_under_test_specification_operator_2 == "" and \
                        device_under_test_specification_operator_3 == "":
                    total_error_band_list.append(float(error_1_value))
                else:
                    if device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value))

                    elif device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value))

                    elif device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_3_value))

                    elif device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value))

                    elif device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_4_value))

                    elif device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_4_value))

                    elif device_under_test_specification_operator_3 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_3_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_4_value))

                    elif device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) - float(error_3_value))

                    elif device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) + float(error_3_value))

                    elif device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) - float(error_3_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) - float(error_4_value))

                    elif device_under_test_specification_operator_3 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_3_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_4_value))

                    elif device_under_test_specification_operator_3 == \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + (float(error_2_value) +
                                                                             float(error_3_value) +
                                                                             float(error_4_value)))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 and \
                            device_under_test_specification_operator_2 and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) -
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 and \
                            device_under_test_specification_operator_2 and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                # //////////////////////////////Pass/Fail Criteria///////////////////////////////////////#

                actual_error_value = ttk.Label(DUTCalibrationWindow,
                                               text=(('{:.' +
                                                      str(reference_resolution) +
                                                      'f}}}').format(round(calculated_measured_difference_value,
                                                                           int(reference_resolution)))).strip("}"),
                                               anchor="n")
                actual_error_value.config(width=measurement_discipline_type_width)
                actual_error_value.grid(row=i + 1, column=5)

                if device_under_test_specification_operator == "<":
                    if calculated_measured_difference < float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == ">":
                    if calculated_measured_difference > float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == u"\u2264":
                    if calculated_measured_difference <= float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == u"\u2265":
                    if calculated_measured_difference >= float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == u"\u00B1":
                    if -float(total_error_band_list[i]) <= calculated_measured_difference <= \
                            float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"

                if judgement_condition == "Pass":
                    pass_fail_condition_list.append(judgement_condition)
                    pass_fail_condition_displayed = ttk.Label(DUTCalibrationWindow, text=judgement_condition,
                                                              anchor="n")
                    pass_fail_condition_displayed.config(width=measurement_discipline_type_width, style='pass.TLabel')
                    pass_fail_condition_displayed.grid(row=i + 1, column=6)
                elif judgement_condition == "Fail":
                    pass_fail_condition_list.append(judgement_condition)
                    pass_fail_condition_displayed = ttk.Label(DUTCalibrationWindow, text=judgement_condition,
                                                              anchor="n")
                    pass_fail_condition_displayed.config(width=measurement_discipline_type_width, style='fail.TLabel')
                    pass_fail_condition_displayed.grid(row=i + 1, column=6)

                i += 1

        # ==================================Transmitter Data==================================#

        elif device_under_test_output_type == "Transmitter":
            i = 0
            for nominal_value_entry, reference_reading_entry, device_under_test_reading_entry in \
                    zip(nominal_input_list, reference_reading_list, device_reading_list):
                x_axis_scaling_list.append(nominal_value_entry.get())
                ideal_transmitter_output = (float(reference_reading_entry.get()) *
                                            (second_device_under_test_full_scale_range / (target_supply_span))) + \
                                           (float(device_under_test_minimum_full_scale_1.get()) -
                                            (float(device_under_test_minimum_full_scale.get()) *
                                             (second_device_under_test_full_scale_range) / (target_supply_span)))
                calculated_measured_difference = (float(device_under_test_reading_entry.get()) -
                                                  float(ideal_transmitter_output))
                calculated_measured_difference_value = round(calculated_measured_difference,
                                                             int(reference_resolution))
                calculated_difference_list.append(calculated_measured_difference_value)

                if device_under_test_specification_type == "Actual Error":
                    error_band_list.append(device_under_test_specification_value)
                    error_1_value = device_under_test_specification_value
                elif device_under_test_specification_type == "% of Full Scale":
                    error_band_list.append(((float(device_under_test_specification_value)) / 100) *
                                           second_device_under_test_full_scale_range)
                    error_1_value = (((float(device_under_test_specification_value)) / 100) *
                                     second_device_under_test_full_scale_range)
                elif device_under_test_specification_type == "% of Reading":
                    error_band_list.append(((float(device_under_test_specification_value)) / 100) *
                                           float(ideal_transmitter_output))
                    error_1_value = (((float(device_under_test_specification_value)) / 100) *
                                     float(ideal_transmitter_output))

                if device_under_test_specification_value_1 is None or device_under_test_specification_value_1 == "" or \
                        device_under_test_specification_value_1 == " ":
                    pass
                else:
                    if device_under_test_specification_type_1 == "Actual Error":
                        error_band_2_list.append(device_under_test_specification_value_1)
                        error_2_value = device_under_test_specification_value_1
                    elif device_under_test_specification_type_1 == "% of Full Scale":
                        error_band_2_list.append(((float(device_under_test_specification_value_1)) / 100) *
                                                 second_device_under_test_full_scale_range)
                        error_2_value = (((float(device_under_test_specification_value_1)) / 100) *
                                         second_device_under_test_full_scale_range)
                    elif device_under_test_specification_type_1 == "% of Reading":
                        error_band_2_list.append(((float(device_under_test_specification_value_1)) / 100) *
                                                 float(ideal_transmitter_output))
                        error_2_value = (((float(device_under_test_specification_value_1)) / 100) *
                                         float(ideal_transmitter_output))

                if device_under_test_specification_value_2 is None or device_under_test_specification_value_2 == "" or \
                        device_under_test_specification_value_2 == " ":
                    pass
                else:
                    if device_under_test_specification_type_2 == "Actual Error":
                        error_band_3_list.append(device_under_test_specification_value_2)
                        error_3_value = device_under_test_specification_value_2
                    elif device_under_test_specification_type_2 == "% of Full Scale":
                        error_band_3_list.append(((float(device_under_test_specification_value_2)) / 100) *
                                                 second_device_under_test_full_scale_range)
                        error_3_value = (((float(device_under_test_specification_value_2)) / 100) *
                                         second_device_under_test_full_scale_range)
                    elif device_under_test_specification_type_2 == "% of Reading":
                        error_band_3_list.append(((float(device_under_test_specification_value_2)) / 100) *
                                                 float(ideal_transmitter_output))
                        error_3_value = (((float(device_under_test_specification_value_2)) / 100) *
                                         float(ideal_transmitter_output))

                if device_under_test_specification_value_3 is None or device_under_test_specification_value_3 == "" or \
                        device_under_test_specification_value_3 == " ":
                    pass
                else:
                    if device_under_test_specification_type_3 == "Actual Error":
                        error_band_4_list.append(device_under_test_specification_value_3)
                        error_4_value = device_under_test_specification_value_3
                    elif device_under_test_specification_type_3 == "% of Full Scale":
                        error_band_4_list.append(((float(device_under_test_specification_value_3)) / 100) *
                                                 second_device_under_test_full_scale_range)
                        error_4_value = (((float(device_under_test_specification_value_3)) / 100) *
                                         second_device_under_test_full_scale_range)
                    elif device_under_test_specification_type_3 == "% of Reading":
                        error_band_4_list.append(((float(device_under_test_specification_value_3)) / 100) *
                                                 float(ideal_transmitter_output))
                        error_4_value = (((float(device_under_test_specification_value_3)) / 100) *
                                         float(ideal_transmitter_output))

                # ////////////////////////Transmitter Error Band Determination////////////////////////#

                if device_under_test_specification_operator_1 == "" and \
                        device_under_test_specification_operator_2 == "" and \
                        device_under_test_specification_operator_3 == "":
                    total_error_band_list.append(float(error_1_value))
                else:
                    if device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value))

                    elif device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value))

                    elif device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_3_value))

                    elif device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value))

                    elif device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_4_value))

                    elif device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_4_value))

                    elif device_under_test_specification_operator_3 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == "" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_3_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_4_value))

                    elif device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) - float(error_3_value))

                    elif device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) + float(error_3_value))

                    elif device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_3 == "+" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) - float(error_3_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) - float(error_4_value))

                    elif device_under_test_specification_operator_3 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) - float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_3 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_3_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) + float(error_4_value))

                    elif device_under_test_specification_operator_3 == \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == "":
                        total_error_band_list.append(float(error_1_value) + float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + (float(error_2_value) +
                                                                             float(error_3_value) + float(error_4_value)))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 and \
                            device_under_test_specification_operator_2 and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) -
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 and \
                            device_under_test_specification_operator_2 and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_3 == "-" and \
                            device_under_test_specification_operator_1 == \
                            device_under_test_specification_operator_2 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == u"\u00B1":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == "+" and \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                    elif device_under_test_specification_operator_1 == "-" and \
                            device_under_test_specification_operator_2 == u"\u00B1" and \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) - float(error_2_value) +
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "-" and \
                            device_under_test_specification_operator_3 == "+":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) -
                                                     float(error_3_value) + float(error_4_value))

                    elif device_under_test_specification_operator_1 == u"\u00B1" and \
                            device_under_test_specification_operator_2 == "+" and \
                            device_under_test_specification_operator_3 == "-":
                        total_error_band_list.append(float(error_1_value) + float(error_2_value) +
                                                     float(error_3_value) - float(error_4_value))

                # //////////////////////////////Pass/Fail Criteria///////////////////////////////////////#

                actual_error_value = ttk.Label(DUTCalibrationWindow,
                                               text=(('{:.' + str(reference_resolution) +
                                                      'f}}}').format(round(calculated_measured_difference_value,
                                                                           int(reference_resolution)))).strip("}"),
                                               anchor="n")
                actual_error_value.config(width=measurement_discipline_type_width)
                actual_error_value.grid(row=i + 1, column=7)

                if device_under_test_specification_operator == "<":
                    if calculated_measured_difference < float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == ">":
                    if calculated_measured_difference > float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == u"\u2264":
                    if calculated_measured_difference <= float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == u"\u2265":
                    if calculated_measured_difference >= float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif device_under_test_specification_operator == u"\u00B1":
                    if -float(total_error_band_list[i]) <= calculated_measured_difference \
                            <= float(total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"

                if judgement_condition == "Pass":
                    transmitter_pass_fail_condition_list.append(judgement_condition)
                    pass_fail_condition_displayed = ttk.Label(DUTCalibrationWindow, text=judgement_condition,
                                                              anchor="n")
                    pass_fail_condition_displayed.config(width=measurement_discipline_type_width, style='pass.TLabel')
                    pass_fail_condition_displayed.grid(row=i + 1, column=8)
                elif judgement_condition == "Fail":
                    transmitter_pass_fail_condition_list.append(judgement_condition)
                    pass_fail_condition_displayed = ttk.Label(DUTCalibrationWindow, text=judgement_condition,
                                                              anchor="n")
                    pass_fail_condition_displayed.config(width=measurement_discipline_type_width, style='fail.TLabel')
                    pass_fail_condition_displayed.grid(row=i + 1, column=8)

                i += 1

        # ==============================Second Dual Output Data===============================#

        i = 0
        if device_under_test_output_type == "Dual":
            for second_device_under_test_reading_entry, second_reference_reading_entry in \
                    zip(second_dut_reading_list, second_reference_reading_list):
                second_x_axis_scaling_list.append(second_device_under_test_reading_entry.get())
                second_measured_difference = float(second_device_under_test_reading_entry.get()) - \
                                             float(second_reference_reading_entry.get())
                second_measured_difference_value = round(second_measured_difference, int(second_reference_resolution))
                second_calculated_difference_list.append(second_measured_difference_value)

                if second_device_under_test_specification_type == "Actual Error":
                    second_error_band_list.append(second_device_under_test_specification_value)
                    second_error_1 = second_device_under_test_specification_value

                elif second_device_under_test_specification_type == "% of Full Scale":
                    second_error_band_list.append(((float(second_device_under_test_specification_value)) / 100) *
                                                  second_device_under_test_full_scale_calculated)
                    second_error_1 = (((float(second_device_under_test_specification_value)) / 100) *
                                      second_device_under_test_full_scale_calculated)

                elif second_device_under_test_specification_type == "% of FS (Range)":
                    second_error_band_list.append(((float(second_device_under_test_specification_value)) / 100) *
                                                  second_device_under_test_full_scale_range)
                    second_error_1 = (((float(second_device_under_test_specification_value)) / 100) *
                                      second_device_under_test_full_scale_range)

                elif device_under_test_specification_type == "% of Reading":
                    second_error_band_list.append(((float(second_device_under_test_specification_value)) / 100) *
                                                  float(second_reference_reading_entry.get()))
                    second_error_1 = (((float(second_device_under_test_specification_value)) / 100) *
                                      float(second_reference_reading_entry.get()))

                if second_device_under_test_specification_value_1 is None or \
                        second_device_under_test_specification_value_1 == "" or \
                        second_device_under_test_specification_value_1 == " ":
                    pass
                else:
                    if second_device_under_test_specification_type_1 == "Actual Error":
                        second_error_band_2_list_.append(second_device_under_test_specification_value_1)
                        second_error_2 = second_device_under_test_specification_value_1

                    elif device_under_test_specification_type_1 == "% of Full Scale":
                        second_error_band_2_list_.append(((float(second_device_under_test_specification_value_1)) /
                                                          100) * second_device_under_test_full_scale_calculated)
                        second_error_2 = (((float(second_device_under_test_specification_value_1)) / 100) *
                                          second_device_under_test_full_scale_calculated)

                    elif device_under_test_specification_type_1 == "% of FS (Range)":
                        second_error_band_2_list_.append(((float(second_device_under_test_specification_value_1)) /
                                                          100) * second_device_under_test_full_scale_range)
                        second_error_2 = (((float(second_device_under_test_specification_value_1)) / 100) *
                                          second_device_under_test_full_scale_range)

                    elif device_under_test_specification_type_1 == "% of Reading":
                        second_error_band_2_list_.append(((float(second_device_under_test_specification_value_1)) / 100)
                                                         * float(second_reference_reading_entry.get()))
                        second_error_2 = (((float(second_device_under_test_specification_value_1)) / 100) *
                                          float(second_reference_reading_entry.get()))

                if second_device_under_test_specification_value_2 is None or \
                        second_device_under_test_specification_value_2 == "" or \
                        second_device_under_test_specification_value_2 == " ":
                    pass
                else:
                    if second_device_under_test_specification_type_2 == "Actual Error":
                        second_error_band_3_list.append(second_device_under_test_specification_value_2)
                        second_error_3 = second_device_under_test_specification_value_2

                    elif second_device_under_test_specification_type_2 == "% of Full Scale":
                        second_error_band_3_list.append(((float(second_device_under_test_specification_value_2)) / 100)
                                                        * second_device_under_test_full_scale_calculated)
                        second_error_3 = (((float(second_device_under_test_specification_value_2)) / 100) *
                                          second_device_under_test_full_scale_calculated)

                    elif second_device_under_test_specification_type_2 == "% of FS (Range)":
                        second_error_band_3_list.append(((float(second_device_under_test_specification_value_2)) / 100)
                                                        * second_device_under_test_full_scale_range)
                        second_error_3 = (((float(second_device_under_test_specification_value_2)) / 100) *
                                          second_device_under_test_full_scale_range)

                    elif device_under_test_specification_type_2 == "% of Reading":
                        second_error_band_3_list.append(((float(second_device_under_test_specification_value_2)) / 100)
                                                        * float(second_reference_reading_entry.get()))
                        second_error_3 = (((float(second_device_under_test_specification_value_2)) / 100) *
                                          float(second_reference_reading_entry.get()))

                if second_device_under_test_specification_value_3 is None or \
                        second_device_under_test_specification_value_3 == "" or \
                        second_device_under_test_specification_value_3 == " ":
                    pass
                else:
                    if second_device_under_test_specification_type_3 == "Actual Error":
                        second_error_band_4_list.append(second_device_under_test_specification_value_3)
                        second_error_4 = second_device_under_test_specification_value_3

                    elif second_device_under_test_specification_type_3 == "% of Full Scale":
                        second_error_band_4_list.append(((float(second_device_under_test_specification_value_3)) / 100)
                                                        * second_device_under_test_full_scale_calculated)
                        second_error_4 = (((float(second_device_under_test_specification_value_3)) / 100) *
                                          second_device_under_test_full_scale_calculated)

                    elif second_device_under_test_specification_type_3 == "% of FS (Range)":
                        second_error_band_4_list.append(((float(second_device_under_test_specification_value_3)) / 100)
                                                        * second_device_under_test_full_scale_range)
                        second_error_4 = (((float(second_device_under_test_specification_value_3)) / 100) *
                                          second_device_under_test_full_scale_range)

                    elif second_device_under_test_specification_type_3 == "% of Reading":
                        second_error_band_4_list.append(((float(second_device_under_test_specification_value_3)) / 100)
                                                        * float(second_reference_reading_entry.get()))
                        second_error_4 = (((float(second_device_under_test_specification_value_3)) / 100) *
                                          float(second_reference_reading_entry.get()))

                # //////////////////////////Dual Output Error Band Determination///////////////////////////#
                if second_device_under_test_specification_operator_1 == "" and \
                        second_device_under_test_specification_operator_2 == "" and \
                        second_device_under_test_specification_operator_3 == "":
                    second_total_error_band_list.append(float(second_error_1))
                else:
                    if second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_2 == "" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_2 == "" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2))

                    elif second_device_under_test_specification_operator_1 == u"\u00B1" and \
                            second_device_under_test_specification_operator_2 == "" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2))

                    elif second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_1 == "" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_3))

                    elif second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_1 == "" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_3))

                    elif second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_1 == "" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_3))

                    elif second_device_under_test_specification_operator_3 == "+" and \
                            second_device_under_test_specification_operator_1 == "" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_4))

                    elif second_device_under_test_specification_operator_3 == "-" and \
                            second_device_under_test_specification_operator_1 == "" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_4))

                    elif second_device_under_test_specification_operator_3 == u"\u00B1" and \
                            second_device_under_test_specification_operator_1 == "" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == "+" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == "+" and \
                            second_device_under_test_specification_operator_1 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_3) +
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_3))

                    elif second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == "-" and \
                            second_device_under_test_specification_operator_1 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_3) -
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_3 == "-" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3))

                    elif second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == "+" and \
                            second_device_under_test_specification_operator_1 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_3) +
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_3 == "+" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) -
                                                            float(second_error_3))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == "-" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) -
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_3 == \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_1 == "":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_3) -
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_3 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == u"\u00B1" and \
                            second_device_under_test_specification_operator_2 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_3 == \
                            second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_1 == "":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_3) +
                                                            float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) + (float(second_error_2) +
                                                                                     float(second_error_3) +
                                                                                     float(second_error_4)))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) -
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 and \
                            second_device_under_test_specification_operator_2 and \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) -
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_1 and \
                            second_device_under_test_specification_operator_2 and \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == u"\u00B1" and \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) -
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == u"\u00B1" and \
                            second_device_under_test_specification_operator_2 == \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_3 == "-" and \
                            second_device_under_test_specification_operator_1 == \
                            second_device_under_test_specification_operator_2 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == u"\u00B1":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "+" and \
                            second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) - float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == "-" and \
                            second_device_under_test_specification_operator_2 == u"\u00B1" and \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) - float(second_error_2) +
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == u"\u00B1" and \
                            second_device_under_test_specification_operator_2 == "-" and \
                            second_device_under_test_specification_operator_3 == "+":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) -
                                                            float(second_error_3) + float(second_error_4))

                    elif second_device_under_test_specification_operator_1 == u"\u00B1" and \
                            second_device_under_test_specification_operator_2 == "+" and \
                            second_device_under_test_specification_operator_3 == "-":
                        second_total_error_band_list.append(float(second_error_1) + float(second_error_2) +
                                                            float(second_error_3) - float(second_error_4))

                # //////////////////////////////Pass/Fail Criteria///////////////////////////////////////#

                actual_error_value = ttk.Label(DUTCalibrationWindow,
                                               text=(('{:.'
                                                      + str(second_reference_resolution)
                                                      + 'f}}}').format(round(second_measured_difference_value,
                                                                             int(second_reference_resolution)))).strip("}"),
                                               anchor="n")
                actual_error_value.config(width=measurement_discipline_type_width)
                actual_error_value.grid(row=i + 1, column=12)

                if second_device_under_test_specification_operator == "<":
                    if second_measured_difference < float(second_total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif second_device_under_test_specification_operator == ">":
                    if second_measured_difference > float(second_total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif second_device_under_test_specification_operator == u"\u2264":
                    if second_measured_difference <= float(second_total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif second_device_under_test_specification_operator == u"\u2265":
                    if second_measured_difference >= float(second_total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"
                elif second_device_under_test_specification_operator == u"\u00B1":
                    if -float(second_total_error_band_list[i]) <= second_measured_difference \
                            <= float(second_total_error_band_list[i]):
                        judgement_condition = "Pass"
                    else:
                        judgement_condition = "Fail"

                if judgement_condition == "Pass":
                    dual_pass_fail_condition_list.append(judgement_condition)
                    pass_fail_condition_displayed = ttk.Label(DUTCalibrationWindow, text=judgement_condition,
                                                              anchor="n")
                    pass_fail_condition_displayed.config(width=measurement_discipline_type_width, style='pass.TLabel')
                    pass_fail_condition_displayed.grid(row=i + 1, column=13)
                elif judgement_condition == "Fail":
                    dual_pass_fail_condition_list.append(judgement_condition)
                    pass_fail_condition_displayed = ttk.Label(DUTCalibrationWindow, text=judgement_condition,
                                                              anchor="n")
                    pass_fail_condition_displayed.config(width=measurement_discipline_type_width, style='fail.TLabel')
                    pass_fail_condition_displayed.grid(row=i + 1, column=13)

                i += 1

        # =====================Conditions to Generate Graph in Tkinter Window=================#

        # matplotlib.use('TkAgg', warn=True)
        #
        # # The following graph is generated for all DUTOutTypes
        # x_axis = np.array(map(float, x_axis_scaling_list))
        # positive_error_band = (np.array(map(float, total_error_band_list)))
        # negative_error_band = (np.array(map(float, total_error_band_list))) * -1
        #
        # fig = plt.figure(figsize=(5.65, 4))
        # canvas = FigureCanvasTkAgg(fig, master=DUTCalibrationWindow)
        # plot_widget = canvas.get_tk_widget()
        #
        # if device_under_test_output_type != "Transmitter":
        #     plot_widget.config(bd=1, relief=SOLID)
        #     plot_widget.grid(row=13, rowspan=5, column=0, columnspan=7, padx=10)
        # elif device_under_test_output_type == "Transmitter":
        #     plot_widget.config(bd=1, relief=SOLID)
        #     plot_widget.grid(row=13, rowspan=5, column=0, columnspan=9, padx=10)
        #
        # plt.clf()  # Clear all graphs drawn in the figure
        # plt.plot(x_axis, calculated_difference_list, 'b*--')
        # plt.plot(x_axis, positive_error_band, 'r-')
        # plt.plot(x_axis, negative_error_band, 'r-')
        # plt.xlabel("Test Points" + "," + units_of_measure)
        #
        # if device_under_test_specification_operator in supplemental_operator_list and \
        #         device_under_test_specification_operator_1 not in supplemental_operator_list and \
        #         device_under_test_specification_operator_2 not in supplemental_operator_list and \
        #         device_under_test_specification_operator_3 not in supplemental_operator_list:
        #     plt.ylabel("DUT Accuracy" + ":" + device_under_test_specification_operator + " " +
        #                device_under_test_specification_value + " " + device_under_test_specification_type)
        #
        # elif device_under_test_specification_operator in supplemental_operator_list and \
        #         device_under_test_specification_operator_1 in supplemental_operator_list and \
        #         device_under_test_specification_operator_2 not in supplemental_operator_list and \
        #         device_under_test_specification_operator_3 not in supplemental_operator_list:
        #     plt.ylabel("DUT Accuracy" + ":" + device_under_test_specification_operator + " " +
        #                device_under_test_specification_value + " " + device_under_test_specification_type + "\n" +
        #                device_under_test_specification_operator_1 + " " + device_under_test_specification_value_1 +
        #                " " + device_under_test_specification_type_1)
        #
        # elif device_under_test_specification_operator in supplemental_operator_list and \
        #         device_under_test_specification_operator_1 in supplemental_operator_list and \
        #         device_under_test_specification_operator_2 in supplemental_operator_list and \
        #         device_under_test_specification_operator_3 not in supplemental_operator_list:
        #     plt.ylabel("DUT Accuracy" + ":" + device_under_test_specification_operator + " " +
        #                device_under_test_specification_value + " " + device_under_test_specification_type + "\n" +
        #                device_under_test_specification_operator_1 + " " + device_under_test_specification_value_1 +
        #                " " + device_under_test_specification_type_1 + "\n" +
        #                device_under_test_specification_operator_2 + " " + device_under_test_specification_value_2 +
        #                " " + device_under_test_specification_type_2)
        #
        # elif device_under_test_specification_operator in supplemental_operator_list and \
        #         device_under_test_specification_operator_1 in supplemental_operator_list and \
        #         device_under_test_specification_operator_2 in supplemental_operator_list and \
        #         device_under_test_specification_operator_3 in supplemental_operator_list:
        #     plt.ylabel("DUT Accuracy" + ":" + device_under_test_specification_operator + " " +
        #                device_under_test_specification_value + " " + device_under_test_specification_type + "\n" +
        #                device_under_test_specification_operator_1 + " " + device_under_test_specification_value_1 +
        #                " " + device_under_test_specification_type_1 + "\n" +
        #                device_under_test_specification_operator_2 + " " + device_under_test_specification_value_2 +
        #                " " + device_under_test_specification_type_2 + "\n" +
        #                device_under_test_specification_operator_3 + " " + device_under_test_specification_value_3 +
        #                " " + device_under_test_specification_type_3)
        #
        # plt.title("Calibration Data")
        #
        # if device_under_test_output_type != "Transmitter":
        #     if device_under_test_calibration_direction == "Ascending":
        #         plt.xlim(float(x_axis_scaling_list[0]) - float(x_axis_scaling_list[1]) / 5.0,
        #                  float(device_under_test_reading_entry.get()) + float(x_axis_scaling_list[1]) / 5.0)
        #
        #     elif device_under_test_calibration_direction == "Descending":
        #         plt.xlim(float(x_axis_scaling_list[0]) + float(x_axis_scaling_list[1]) / 5.0,
        #                  float(device_under_test_reading_entry.get()) - float(x_axis_scaling_list[1]) / 5.0)
        #
        # elif device_under_test_output_type == "Transmitter":
        #     if device_under_test_calibration_direction == "Ascending":
        #         plt.xlim(float(x_axis_scaling_list[0]) - float(x_axis_scaling_list[1]) / 5.0,
        #                  float(x_axis_scaling_list[-1]) + float(x_axis_scaling_list[1]) / 5.0)
        #
        #     elif device_under_test_calibration_direction == "Descending":
        #         plt.xlim(float(x_axis_scaling_list[0]) + float(x_axis_scaling_list[1]) / 5.0,
        #                  float(x_axis_scaling_list[-1]) - float(x_axis_scaling_list[1]) / 5.0)
        #
        # plt.ylim(float(total_error_band_list[-1]) * -2.0, float(total_error_band_list[-1]) * 2.0)
        # fig.tight_layout()
        # fig.canvas.draw()
        #
        # # This graph is only generated if DUTOutType is set to Dual
        # if device_under_test_output_type == "Dual":
        #     second_x_axis = np.array(map(float, second_x_axis_scaling_list))
        #     second_positive_error_band = (np.array(map(float, second_total_error_band_list)))
        #     second_negative_error_band = (np.array(map(float, second_total_error_band_list))) * -1
        #
        #     fig2 = plt.figure(figsize=(5.65, 4))
        #     canvas = FigureCanvasTkAgg(fig2, master=DUTCalibrationWindow)
        #     plot_widget = canvas.get_tk_widget()
        #     plot_widget.config(bd=1, relief=SOLID)
        #     plot_widget.grid(row=13, rowspan=5, column=7, columnspan=7, padx=10)
        #
        #     plt.clf()  # Clear all graphs drawn in the figure
        #     plt.plot(second_x_axis, second_calculated_difference_list, 'b*--')
        #     plt.plot(second_x_axis, second_positive_error_band, 'r-')
        #     plt.plot(second_x_axis, second_negative_error_band, 'r-')
        #     plt.xlabel("Test Points" + "," + second_units_of_measure)
        #
        #     if second_device_under_test_specification_operator in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_1 not in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_2 not in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_3 not in supplemental_operator_list:
        #         plt.ylabel("DUT Accuracy" + ":" + second_device_under_test_specification_operator + " " +
        #                    second_device_under_test_specification_value + " " +
        #                    second_device_under_test_specification_type)
        #
        #     elif second_device_under_test_specification_operator in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_1 in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_2 not in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_3 not in supplemental_operator_list:
        #         plt.ylabel("DUT Accuracy" + ":" + second_device_under_test_specification_operator + " " +
        #                    second_device_under_test_specification_value + " " +
        #                    second_device_under_test_specification_type + "\n" +
        #                    second_device_under_test_specification_operator_1 + " " +
        #                    second_device_under_test_specification_value_1 + " " +
        #                    second_device_under_test_specification_type_1)
        #
        #     elif second_device_under_test_specification_operator in supplemental_operator_list and \
        #             device_under_test_specification_operator_1 in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_2 in supplemental_operator_list and \
        #             device_under_test_specification_operator_3 not in supplemental_operator_list:
        #         plt.ylabel("DUT Accuracy" + ":" + second_device_under_test_specification_operator + " " +
        #                    second_device_under_test_specification_value + " " +
        #                    second_device_under_test_specification_type + "\n" +
        #                    second_device_under_test_specification_operator_1 + " " +
        #                    second_device_under_test_specification_value_1 + " " +
        #                    second_device_under_test_specification_type_1 + "\n" +
        #                    second_device_under_test_specification_operator_2 + " " +
        #                    second_device_under_test_specification_value_2 + " " +
        #                    second_device_under_test_specification_type_2)
        #
        #     elif second_device_under_test_specification_operator in supplemental_operator_list and \
        #             device_under_test_specification_operator_1 in supplemental_operator_list and \
        #             second_device_under_test_specification_operator_2 in supplemental_operator_list and \
        #             device_under_test_specification_operator_3 in supplemental_operator_list:
        #         plt.ylabel("DUT Accuracy" + ":" + second_device_under_test_specification_operator + " " +
        #                    second_device_under_test_specification_value + " " +
        #                    second_device_under_test_specification_type + "\n" +
        #                    second_device_under_test_specification_operator_1 + " " +
        #                    second_device_under_test_specification_value_1 + " " +
        #                    second_device_under_test_specification_type_1 + "\n" +
        #                    second_device_under_test_specification_operator_2 + " " +
        #                    second_device_under_test_specification_value_2 + " " +
        #                    second_device_under_test_specification_type_2 + "\n" +
        #                    second_device_under_test_specification_operator_3 + " " +
        #                    second_device_under_test_specification_value_3 + " " +
        #                    second_device_under_test_specification_type_3)
        #
        #     plt.title("Calibration Data")
        #     if device_under_test_calibration_direction == "Ascending":
        #         plt.xlim(float(second_x_axis_scaling_list[0]) - float(second_x_axis_scaling_list[1]) / 5.0,
        #                  float(second_device_under_test_reading_entry.get()) +
        #                  float(second_x_axis_scaling_list[1]) / 5.0)
        #
        #     elif device_under_test_calibration_direction == "Descending":
        #         plt.xlim(float(second_x_axis_scaling_list[0]) + float(second_x_axis_scaling_list[1]) / 5.0,
        #                  float(second_device_under_test_reading_entry.get()) -
        #                  float(second_x_axis_scaling_list[1]) / 5.0)
        #
        #     plt.ylim(float(second_total_error_band_list[-1]) * -2.0, float(second_total_error_band_list[-1]) * 2.0)
        #
        #     fig2.tight_layout()
        #     fig2.canvas.draw()

    # -----------------------------------------------------------------------#
    # Command to Ask User if They are Ready to Generate Certificate of Calibration
    def complete_calibration_process_step_one(self):

        certificate_of_calibration_query = tm.askyesno("Complete Calibration", "Would you like to generate a \
Certificate of Calibration for the device under test?")
        if certificate_of_calibration_query is True:
            self.complete_calibration_process_step_two(DUTCalibrationWindow)
        elif certificate_of_calibration_query is False:
            self.calculate_calibration_measurement_error()

    # -----------------------------------------------------------------------#
    # Command to Ask User for Notes and to Sign Off on Certificate of Calibration
    def complete_calibration_process_step_two(self, window):
        window.withdraw()

        global CertNoteSignOff, technician_notes, personnel_list, personnel_selection, \
            personnel_selection, btn_approve_and_create, calibration_time, pgrbar_approve_and_create

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        aph = AppHelpWindows()

        time_spent_on_calibration = StringVar()

        # ........................Main Window Properties.........................#

        CertNoteSignOff = Toplevel()
        CertNoteSignOff.title("Certificate of Calibration - Signature Required")
        CertNoteSignOff.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 370
        height = 300
        screen_width = CertNoteSignOff.winfo_screenwidth()
        screen_height = CertNoteSignOff.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        CertNoteSignOff.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CertNoteSignOff.focus_force()
        CertNoteSignOff.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CertNoteSignOff))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menu_bar = MenuBar(CertNoteSignOff, width, height, x, y)
        menu_bar.add_menu("Help", commands=[("Help", lambda: aph.technician_notes_and_signature_help())])

        # ..........................Frame Creation...............................#

        note_and_signature_frame = LabelFrame(CertNoteSignOff, text="Customer Notes & User Signature",
                                              relief=SOLID, bd=1, labelanchor="n")
        note_and_signature_frame.grid(row=0, column=0, rowspan=5, columnspan=5, padx=8, pady=8)

        # ........................Labels and Entries.............................#

        # Ask for Notes to be Displayed on Calibration Certificate
        lbl_certificate_notes = ttk.Label(note_and_signature_frame, text="Certificate Notes:",
                                          font=('arial', 11), anchor="n")
        lbl_certificate_notes.grid(row=1, padx=5, pady=5)
        technician_notes = Text(note_and_signature_frame, wrap=WORD, height=5, width=25, font=('arial', 11))
        technician_notes.grid(row=1, column=1, rowspan=3)

        # Ask for Name for Signature and Job Title on Certificate
        personnel_list = [" ", "Jason Berry", "Steve Evert", "Robert Maldonado", "Randall Massner", "Dave Niezgodzki",
                          "Roger Shumaker", "Marco Suarez", "Jeff Woodruff", "Frank Engel"]
        lbl_personnel_name = ttk.Label(note_and_signature_frame, text="Name:", font=('arial', 11), anchor="n")
        lbl_personnel_name.grid(row=5, padx=5)
        personnel_selection = ttk.Combobox(note_and_signature_frame, values=[" ", "Jason Berry", "Steve Evert",
                                                                             "Robert Maldonado", "Randall Massner",
                                                                             "Dave Niezgodzki", "Roger Shumaker",
                                                                             "Marco Suarez", "Jeff Woodruff"])
        acc.always_active_style(personnel_selection)
        personnel_selection.config(state="active", width=21)
        personnel_selection.focus()
        personnel_selection.grid(row=5, column=1, padx=5, pady=8)

        # This label is a dummy label that does nothing but help format the window
        dummy = Label(note_and_signature_frame)
        dummy.grid(row=5, column=5)
        dummy.config(width=2)

        # Ask for Time Spent on Calibration
        lbl_time_spent = ttk.Label(note_and_signature_frame, text="Time Spent:", font=('arial', 11), anchor="n")
        lbl_time_spent.grid(row=6, padx=5)
        calibration_time = ttk.Entry(note_and_signature_frame, textvariable=time_spent_on_calibration,
                                     font=('arial', 11))
        calibration_time.grid(row=6, column=1, pady=8)
        calibration_time.config(width=18)

        # ...............................Buttons..................................#

        btn_approve_and_create = ttk.Button(note_and_signature_frame, text="Sign and Generate",
                                            command=lambda: self.verify_certificate_signature_input())
        btn_approve_and_create.bind("<Return>", lambda event: self.verify_certificate_signature_input())
        btn_approve_and_create.grid(row=7, column=0, columnspan=2, pady=5)
        btn_approve_and_create.config(width=25)

        pgrbar_approve_and_create = ttk.Progressbar(note_and_signature_frame, orient='horizontal',
                                                    length=100, mode='determinate')
        pgrbar_approve_and_create.grid(row=8, column=0, columnspan=2, padx=8, pady=8)
        pgrbar_approve_and_create['value'] = 0

    # -----------------------------------------------------------------------#

    # Command to Verify Signature/User Has Been Selected for Generating Certificate
    def verify_certificate_signature_input(self):

        if personnel_selection.get() == personnel_list[0]:
            tm.showerror("Signature Required", "Please select digital signature from drop down list provided.")
        else:
            self.job_title_and_personnel_signature()
            from LIMSCertDBase import AppCertificateDatabase
            apc = AppCertificateDatabase()
            try:
                year = datetime.today().year
                excel_database = open("\\\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls" %(year, year))
                if excel_database.closed is False:
                    excel_database.close()
                    apc.certificate_number_fill_in()
                    self.generate_certificate_of_calibration_report()
            except IOError as e:
                btn_device_specification.config(cursor="arrow")
                tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------#

    # Command to Pull and Store Information Correlated to Digital Signature Selection
    def job_title_and_personnel_signature(self):

        if personnel_selection.get() == personnel_list[1]:
            LIMSVarConfig.certificate_technician_job_title = 'Engineering Laboratory Supervisor'
        elif personnel_selection.get() == personnel_list[2]:
            LIMSVarConfig.certificate_technician_job_title = 'Engineering Technician II'
        elif personnel_selection.get() == personnel_list[3]:
            LIMSVarConfig.certificate_technician_job_title = 'Engineering Lab. Tech., Generalist'
        elif personnel_selection.get() == personnel_list[4]:
            LIMSVarConfig.certificate_technician_job_title = 'Sr. Electrical Engineering Technician'
        elif personnel_selection.get() == personnel_list[5]:
            LIMSVarConfig.certificate_technician_job_title = 'Sr. Engineering Technician'
        elif personnel_selection.get() == personnel_list[6]:
            LIMSVarConfig.certificate_technician_job_title = 'Group Leader / Sr. Elec. Eng. Tech.'
        elif personnel_selection.get() == personnel_list[7]:
            LIMSVarConfig.certificate_technician_job_title = 'Engineering Technician II'
        elif personnel_selection.get() == personnel_list[8]:
            LIMSVarConfig.certificate_technician_job_title = 'Sr. Electrical Engineering Technician'
        elif personnel_selection.get() == personnel_list[9]:
            LIMSVarConfig.certificate_technician_job_title = 'Engineering Laboratory Technician'
        else:
            LIMSVarConfig.certificate_technician_job_title = ""
        LIMSVarConfig.certificate_technician_notes = technician_notes.get('1.0', 'end-1c')

        if personnel_selection.get() == personnel_list[0] or personnel_selection.get() == "":
            LIMSVarConfig.certificate_technician_name = ""
            LIMSVarConfig.certificate_technician_signature = ""
        else:
            LIMSVarConfig.certificate_technician_name = personnel_selection.get()
            LIMSVarConfig.certificate_technician_signature = personnel_selection.get()

        LIMSVarConfig.calibration_time_helper = calibration_time.get()

    # -----------------------------------------------------------------------#

    # Command to Export All Information Entered in Program and Print on
    # Certificate Form Window for Review Prior to Printing
    def generate_certificate_of_calibration_report(self):
        global f, ws, ws

        LIMSVarConfig.certificate_of_calibration_entry_list = []

        btn_approve_and_create.config(cursor="watch")
        pgrbar_approve_and_create['value'] = 10
        CertNoteSignOff.update_idletasks()

        from LIMSEnvirCond import AppEnvironmentalConditions
        aec = AppEnvironmentalConditions()
        aec.main_lab_environmental_conditions_query()
        aec.flow_lab_conditions_query()

        pgrbar_approve_and_create['value'] = 20
        CertNoteSignOff.update_idletasks()

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        wb = excel.Workbooks.Open(
            '\\\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\LIMS CCT Files\\Templates\\Dwyer LIMS Cert Form.xlsx')

        # ............................Sheet Selection............................#

        if device_under_test_output_type == "Single":
            ws = wb.Sheets("Single, Plain, AFAL")
        elif device_under_test_output_type == "Dual":
            ws = wb.Sheets("Dual, Plain, AFAL")
        elif device_under_test_output_type == "Transmitter":
            ws = wb.Sheets("Transmitter, Plain, AFAL")

        # /////////////////////////////Fill Out Form//////////////////////////////#

        # ..........................Customer Information..........................#

        if LIMSVarConfig.customer_selection_check == int(0):
            ws.Range('C9').Value = LIMSVarConfig.external_customer_name
            ws.Range('C10').Value = LIMSVarConfig.external_customer_address
            ws.Range('C11').Value = LIMSVarConfig.external_customer_address_1
            ws.Range('C12').Value = LIMSVarConfig.external_customer_address_2
            if ws.Range('C12').Value is None:
                ws.Range('C12').Value = LIMSVarConfig.external_customer_city + " " + \
                                        LIMSVarConfig.external_customer_state_country_zip
                ws.Range('C13').Value = " "
            else:
                ws.Range('C13').Value = LIMSVarConfig.external_customer_city + " " + \
                                        LIMSVarConfig.external_customer_state_country_zip
            if ws.Range('C11').Value is None:
                ws.Range('C11').Value = LIMSVarConfig.external_customer_city + " " + \
                                        LIMSVarConfig.external_customer_state_country_zip
                ws.Range('C12').Value = " "
                ws.Range('C13').Value = " "
            ws.Range('C14').Value = LIMSVarConfig.external_customer_po
        else:
            ws.Range('C9').Value = LIMSVarConfig.internal_customer_name
            ws.Range('C10').Value = LIMSVarConfig.internal_customer_location
            ws.Range('C11').Value = LIMSVarConfig.internal_customer_address_displayed
            ws.Range('C12').Value = LIMSVarConfig.internal_customer_city_displayed
            ws.Range('C13').Value = LIMSVarConfig.internal_customer_state_city_zip_displayed
            ws.Range('C14').Value = "-"

        pgrbar_approve_and_create['value'] = 30
        CertNoteSignOff.update_idletasks()

        # ...................Certificate of Calibration Information...............#

        ws.Range('I9').Value = LIMSVarConfig.certificate_of_calibration_number

        if LIMSVarConfig.customer_selection_check == int(0):
            ws.Range('I10').Value = LIMSVarConfig.external_customer_sales_order_number_helper
            ws.Range('I11').Value = LIMSVarConfig.external_customer_rma_number
        else:
            ws.Range('I10').Value = "-"
            ws.Range('I11').Value = "-"

        ws.Range('I13').Value = LIMSVarConfig.calibration_date_helper
        ws.Range('I14').Value = LIMSVarConfig.calibration_due_date_helper

        pgrbar_approve_and_create['value'] = 40
        CertNoteSignOff.update_idletasks()

        # ......................Calibration Standard Information..................#

        ws.Range('A21').Value = LIMSVarConfig.calibration_standard_equipment_description
        ws.Range('A22').Value = LIMSVarConfig.calibration_standard_equipment_description_2
        ws.Range('A23').Value = LIMSVarConfig.calibration_standard_equipment_description_3
        ws.Range('A24').Value = LIMSVarConfig.calibration_standard_equipment_description_4
        ws.Range('A25').Value = LIMSVarConfig.calibration_standard_equipment_description_5
        ws.Range('A26').Value = LIMSVarConfig.calibration_standard_equipment_description_6
        ws.Range('A27').Value = LIMSVarConfig.calibration_standard_equipment_description_7
        ws.Range('A28').Value = LIMSVarConfig.calibration_standard_equipment_description_8
        ws.Range('E21').Value = LIMSVarConfig.calibration_standard_serial_number
        ws.Range('E22').Value = LIMSVarConfig.calibration_standard_serial_number_2
        ws.Range('E23').Value = LIMSVarConfig.calibration_standard_serial_number_3
        ws.Range('E24').Value = LIMSVarConfig.calibration_standard_serial_number_4
        ws.Range('E25').Value = LIMSVarConfig.calibration_standard_serial_number_5
        ws.Range('E26').Value = LIMSVarConfig.calibration_standard_serial_number_6
        ws.Range('E27').Value = LIMSVarConfig.calibration_standard_serial_number_7
        ws.Range('E28').Value = LIMSVarConfig.calibration_standard_serial_number_8
        ws.Range('G21').Value = calibration_standard_1.get()
        ws.Range('G22').Value = calibration_standard_2.get()
        ws.Range('G23').Value = calibration_standard_3.get()
        ws.Range('G24').Value = calibration_standard_4.get()
        ws.Range('G25').Value = calibration_standard_5.get()
        ws.Range('G26').Value = calibration_standard_6.get()
        ws.Range('G27').Value = calibration_standard_7.get()
        ws.Range('G28').Value = calibration_standard_8.get()
        ws.Range('I21').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date
        ws.Range('I22').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_2
        ws.Range('I23').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_3
        ws.Range('I24').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_4
        ws.Range('I25').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_5
        ws.Range('I26').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_6
        ws.Range('I27').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_7
        ws.Range('I28').Value = LIMSVarConfig.calibration_standard_equipment_calibration_date_8
        ws.Range('K21').Value = LIMSVarConfig.calibration_standard_equipment_due_date
        ws.Range('K22').Value = LIMSVarConfig.calibration_standard_equipment_due_date_2
        ws.Range('K23').Value = LIMSVarConfig.calibration_standard_equipment_due_date_3
        ws.Range('K24').Value = LIMSVarConfig.calibration_standard_equipment_due_date_4
        ws.Range('K25').Value = LIMSVarConfig.calibration_standard_equipment_due_date_5
        ws.Range('K26').Value = LIMSVarConfig.calibration_standard_equipment_due_date_6
        ws.Range('K27').Value = LIMSVarConfig.calibration_standard_equipment_due_date_7
        ws.Range('K28').Value = LIMSVarConfig.calibration_standard_equipment_due_date_8

        pgrbar_approve_and_create['value'] = 50
        CertNoteSignOff.update_idletasks()

        # ..........................Instrument Information.......................#

        ws.Range('K32').Value = condition_of_dut
        ws.Range('C33').Value = LIMSVarConfig.device_identification_number_helper
        ws.Range('G33').Value = device_customer_identification_number.get()
        ws.Range('K33').Value = LIMSVarConfig.device_date_code_helper
        ws.Range('C34').Value = dut_model_number.get()
        ws.Range('G34').Value = device_under_test_measurement_type
        ws.Range('K34').Value = units_of_measure
        ws.Range('C35').Value = device_under_test_minimum
        ws.Range('G35').Value = device_under_test_maximum
        ws.Range('K35').Value = device_under_test_measurement_resolution
        if device_under_test_output_type != "Single":
            ws.Range('G37').Value = second_device_under_test_measurement_type
            ws.Range('K37').Value = second_units_of_measure
            ws.Range('C38').Value = second_device_under_test_minimum
            ws.Range('G38').Value = second_device_under_test_maximum
            if device_under_test_output_type != "Transmitter":
                ws.Range('K38').Value = second_device_under_test_measurement_resolution
            else:
                ws.Range('K38').Value = "-"

        pgrbar_approve_and_create['value'] = 60
        CertNoteSignOff.update_idletasks()

        # ......................Environmental Conditions..........................#

        if device_under_test_output_type == "Single":
            if device_under_test_measurement_type != "Flow" and device_under_test_measurement_type != "Velocity":
                ws.Range('C40').Value = LIMSVarConfig.temperature_0.strip("F")
                ws.Range('F40').Value = LIMSVarConfig.humidity_0.strip("%RH")
                ws.Range('I40').Value = LIMSVarConfig.pressure_0.strip("inHg")
                ws.Range('L40').Value = LIMSVarConfig.dew_point_0.strip("F")
            else:
                ws.Range('C40').Value = LIMSVarConfig.temperature_1.strip("F")
                ws.Range('F40').Value = LIMSVarConfig.humidity_1.strip("%RH")
                ws.Range('I40').Value = LIMSVarConfig.pressure_1.strip("inHg")
                ws.Range('L40').Value = LIMSVarConfig.dew_point_1.strip("F")
        elif device_under_test_output_type != "Single":
            if device_under_test_measurement_type != "Flow" and device_under_test_measurement_type != "Velocity" and \
                    second_device_under_test_measurement_type != "Flow" and \
                    second_device_under_test_measurement_type != "Velocity":
                ws.Range('C43').Value = LIMSVarConfig.temperature_0.strip("F")
                ws.Range('F43').Value = LIMSVarConfig.humidity_0.strip("%RH")
                ws.Range('I43').Value = LIMSVarConfig.pressure_0.strip("inHg")
                ws.Range('L43').Value = LIMSVarConfig.dew_point_0.strip("F")
            else:
                ws.Range('C43').Value = LIMSVarConfig.temperature_1.strip("F")
                ws.Range('F43').Value = LIMSVarConfig.humidity_1.strip("%RH")
                ws.Range('I43').Value = LIMSVarConfig.pressure_1.strip("inHg")
                ws.Range('L43').Value = LIMSVarConfig.dew_point_1.strip("F")

        pgrbar_approve_and_create['value'] = 70
        CertNoteSignOff.update_idletasks()

        # .........Notes, Calibrated By, Signature, Job Title & Procedure.........#

        if device_under_test_output_type == "Single":
            ws.Range('A44').Value = LIMSVarConfig.certificate_technician_notes
            ws.Range('C54').Value = LIMSVarConfig.certificate_technician_name
            ws.Range('C55').Value = LIMSVarConfig.certificate_technician_signature
            ws.Range('C56').Value = LIMSVarConfig.certificate_technician_job_title
            ws.Range('I54').Value = "EP-00055-B"
        elif device_under_test_output_type != "Single":
            ws.Range('A47').Value = LIMSVarConfig.certificate_technician_notes
            ws.Range('C57').Value = LIMSVarConfig.certificate_technician_name
            ws.Range('C58').Value = LIMSVarConfig.certificate_technician_signature
            ws.Range('C59').Value = LIMSVarConfig.certificate_technician_job_title
            ws.Range('I57').Value = "EP-00055-B"

        pgrbar_approve_and_create['value'] = 80
        CertNoteSignOff.update_idletasks()

        # ..........................Calibration Data .............................#

        if device_under_test_output_type != "Transmitter":
            row = 79
            if LIMSVarConfig.imported_data_checker == int(0):
                for dut_reading, ref_reading in zip(device_reading_list, reference_reading_list):
                    ws.Cells(row, 2).Value = dut_reading.get()
                    ws.Cells(row, 4).Value = ref_reading.get()
                    ws.Cells(row, 6).Value = units_of_measure
                    row += 1
            else:
                for entry1, entry2 in \
                        zip(LIMSVarConfig.imported_device_reading_list, LIMSVarConfig.imported_reference_reading_list):
                    ws.Cells(row, 2).Value = str(entry1)
                    ws.Cells(row, 4).Value = str(entry2)
                    ws.Cells(row, 6).Value = units_of_measure
                    row += 1
                if LIMSVarConfig.second_imported_data_checker == int(1):
                    if device_under_test_output_type == "Single":
                        row_2 = 139
                        for second_entry1, second_entry2 in \
                                zip(LIMSVarConfig.second_imported_device_reading_list,
                                    LIMSVarConfig.second_imported_reference_reading_list):
                            ws.Cells(row_2, 2).Value = str(second_entry1)
                            ws.Cells(row_2, 4).Value = str(second_entry2)
                            ws.Cells(row_2, 6).Value = units_of_measure
                            row_2 += 1
                    else:
                        self.__init__()
                else:
                    self.__init__()
            row = 79
            if LIMSVarConfig.imported_data_checker == int(0):
                for entry in total_error_band_list:
                    ws.Cells(row, 9).Value = entry
                    row += 1
            else:
                for entry3 in LIMSVarConfig.imported_total_error_band_list:
                    ws.Cells(row, 9).Value = entry3
                    row += 1
                if LIMSVarConfig.second_imported_data_checker == int(1):
                    if device_under_test_output_type == "Single":
                        row_2 = 139
                        for entry4 in LIMSVarConfig.second_imported_total_error_band_list:
                            ws.Cells(row_2, 9).Value = entry4
                            row_2 += 1
                    else:
                        self.__init__()
                else:
                    self.__init__()

        elif device_under_test_output_type == "Transmitter":
            row = 79
            for nominal_value, ref_reading, dut_reading in zip(nominal_input_list, reference_reading_list,
                                                               device_reading_list):
                ws.Cells(row, 1).Value = nominal_value.get()
                ws.Cells(row, 3).Value = ref_reading.get()
                ws.Cells(row, 5).Value = units_of_measure
                ws.Cells(row, 6).Value = dut_reading.get()
                row += 1
            row = 79
            for entry in total_error_band_list:
                ws.Cells(row, 10).Value = entry
                row += 1
        if device_under_test_output_type == "Dual":
            row = 79
            for dut_reading, ref_reading in zip(device_reading_list, reference_reading_list):
                ws.Cells(row, 2).Value = dut_reading.get()
                ws.Cells(row, 4).Value = ref_reading.get()
                ws.Cells(row, 6).Value = units_of_measure
                row += 1
            row = 89
            for SecondDUTReading, SecondREFReading in zip(second_dut_reading_list, second_reference_reading_list):
                ws.Cells(row, 2).Value = SecondDUTReading.get()
                ws.Cells(row, 4).Value = SecondREFReading.get()
                ws.Cells(row, 6).Value = second_units_of_measure
                row += 1
            row = 89
            for entry in second_total_error_band_list:
                ws.Cells(row, 9).Value = entry
                row += 1

        if LIMSVarConfig.second_imported_data_checker == int(1) and device_under_test_output_type == "Single":
            ws.Range('B77').Value = "Calibration Data: As Found"
            ws.Range('B137').Value = "Calibration Data: As Left"

        pgrbar_approve_and_create['value'] = 90
        CertNoteSignOff.update_idletasks()

        # ......................Save and Close File...............................#
        year = datetime.today().year
        if device_under_test_output_type == "Single":
            if LIMSVarConfig.imported_data_checker == int(0):
                if 'Fail' in pass_fail_condition_list:
                    ws.Range('I9').Value = LIMSVarConfig.certificate_of_calibration_number + 'F'
                    ws.Range('I14').Value = "-"
                    f = '\\\\\\BDC5\\certdbase\\%s\\%sF.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                    wb.SaveAs(f)

                    from LIMSCertDBase import AppCertificateDatabase
                    acd = AppCertificateDatabase()
                    acd.certificate_failure()

                elif 'Fail' not in pass_fail_condition_list:
                    f = '\\\\\\BDC5\\certdbase\\%s\\%s.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                    wb.SaveAs(f)
            else:
                if "Fail" in LIMSVarConfig.imported_pass_fail_list:
                    ws.Range('I9').Value = LIMSVarConfig.certificate_of_calibration_number + 'F'
                    ws.Range('I14').Value = "-"

                    f = '\\\\\\BDC5\\certdbase\\%s\\%sF.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                    wb.SaveAs(f)

                    from LIMSCertDBase import AppCertificateDatabase
                    acd = AppCertificateDatabase()
                    acd.certificate_failure()

                elif 'Fail' not in LIMSVarConfig.imported_pass_fail_list:
                    f = '\\\\\\BDC5\\certdbase\\%s\\%s.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                    wb.SaveAs(f)

        elif device_under_test_output_type == "Dual":
            if 'Fail' in dual_pass_fail_condition_list:
                ws.Range('I9').Value = LIMSVarConfig.certificate_of_calibration_number + 'F'
                ws.Range('I14').Value = "-"

                f = '\\\\\\BDC5\\certdbase\\%s\\%sF.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                wb.SaveAs(f)

                from LIMSCertDBase import AppCertificateDatabase
                acd = AppCertificateDatabase()
                acd.certificate_failure()

            elif 'Fail' not in dual_pass_fail_condition_list:
                f = '\\\\\\BDC5\\certdbase\\%s\\%s.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                wb.SaveAs(f)

        elif device_under_test_output_type == "Transmitter":
            if 'Fail' in transmitter_pass_fail_condition_list:
                ws.Range('I9').Value = LIMSVarConfig.certificate_of_calibration_number + 'F'
                ws.Range('I14').Value = "-"

                f = '\\\\\\BDC5\\certdbase\\%s\\%sF.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                wb.SaveAs(f)

                from LIMSCertDBase import AppCertificateDatabase
                acd = AppCertificateDatabase()
                acd.certificate_failure()

            elif 'Fail' not in transmitter_pass_fail_condition_list:
                f = '\\\\\\BDC5\\certdbase\\%s\\%s.xlsx' %(year, LIMSVarConfig.certificate_of_calibration_number)
                wb.SaveAs(f)
        wb.Close(True)

        pgrbar_approve_and_create['value'] = 100
        CertNoteSignOff.update_idletasks()

        # Open Certificate of Calibration
        os.startfile(f)

        btn_approve_and_create.config(cursor="arrow")
        CertNoteSignOff.withdraw()
        self.additional_certificate_of_calibration_query()

    # -----------------------------------------------------------------------#

    # Command to Ask User if They Want to Generate Another Certificate of Calibration
    def additional_certificate_of_calibration_query(self):
        global additional_certificate_of_calibration_response
        additional_certificate_of_calibration_response = tm.askyesno("Generate Another Certificate?", "Would you like \
to generate another certificate of calibration? \n\n\
NOTE: YOU SHOULD ONLY CLICK THE 'YES' BUTTON IF YOU PLAN TO CALIBRATE ANOTHER DEVICE WITH THE SAME EXACT INFORMATION \
(i.e. Accuracy, Range) FOR THE EXACT SAME CUSTOMER. If no, click the 'No' button to be taken back to the Main \
Menu. \n\n\
Be sure to review your newly created Certificate of Calibration.")
        if additional_certificate_of_calibration_response is True:
            LIMSVarConfig.certificate_of_calibration_number = ""
            LIMSVarConfig.device_serial_number = ""
            self.additional_calibration_certificate_creation()
        elif additional_certificate_of_calibration_response is False:
            LIMSVarConfig.clear_all_certificate_of_calibration_variables()
            from LIMSHomeWindow import AppHomeWindow
            aph = AppHomeWindow()
            aph.home_window()

    # -----------------------------------------------------------------------#

    # Command to Ask User for Input Regarding Additional DUT Calibration
    def additional_calibration_certificate_creation(self):

        global AdditionalCalCertDUTDetails, new_additional_certificate_number, new_additional_serial_number, \
            device_customer_identification_number, customer_device_under_test_instrument_number_2, \
            device_under_test_date_code_2, device_under_test_date_code_2, calibration_date_2, \
            device_under_test_calibration_date_2, due_date_2, device_under_test_due_date_2, \
            btn_import_new_certificate_number, btn_import_new_certificate_and_serial_number, \
            device_under_test_instrument_id_number_2, additional_dut_details_frame

        device_under_test_instrument_id_number_2 = StringVar()
        customer_device_under_test_instrument_number_2 = StringVar()
        device_under_test_date_code_2 = StringVar()
        device_under_test_calibration_date_2 = StringVar()
        device_under_test_due_date_2 = StringVar()

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        aph = AppHelpWindows()

        # ........................Main Window Properties.........................#

        AdditionalCalCertDUTDetails = Toplevel()
        AdditionalCalCertDUTDetails.title("Additional Certificate of Calibration Creation")
        AdditionalCalCertDUTDetails.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 545
        height = 285
        screen_width = AdditionalCalCertDUTDetails.winfo_screenwidth()
        screen_height = AdditionalCalCertDUTDetails.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        AdditionalCalCertDUTDetails.geometry("%dx%d+%d+%d" % (width, height, x, y))
        AdditionalCalCertDUTDetails.focus_force()
        AdditionalCalCertDUTDetails.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(AdditionalCalCertDUTDetails))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(AdditionalCalCertDUTDetails, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: self.exit_additional_calibration()),
                                           ("Logout", lambda: acc.software_signout(AdditionalCalCertDUTDetails)),
                                           ("Quit", lambda: acc.software_close(AdditionalCalCertDUTDetails))])
        menubar.add_menu("Help", commands=[("Help", lambda: aph.additional_certificate_generation_help())])

        # ...........................Frame Creation..............................#

        additional_dut_details_frame = LabelFrame(AdditionalCalCertDUTDetails, text="Instrument Information",
                                                  relief=SOLID, bd=1, labelanchor="n")
        additional_dut_details_frame.grid(row=0, column=0, rowspan=8, columnspan=4, padx=8, pady=5)

        # .........................Labels and Entries............................#

        # Ask for Certificate of Calibration Number
        lbl_additional_certificate_number = ttk.Label(additional_dut_details_frame, text="Certificate Number:",
                                                      font=('arial', 12), anchor="n")
        lbl_additional_certificate_number.grid(row=1, pady=5)
        new_additional_certificate_number = ttk.Label(additional_dut_details_frame,
                                                      text=LIMSVarConfig.certificate_of_calibration_number, font=12)
        new_additional_certificate_number.grid(row=1, column=1)
        new_additional_certificate_number.config(width=15)

        # Ask for Device Serial Number
        lbl_additional_serial_number = ttk.Label(additional_dut_details_frame, text="Serial Number:",
                                                 font=('arial', 12), anchor="n")
        lbl_additional_serial_number.grid(row=2, pady=5)
        new_additional_serial_number = ttk.Entry(additional_dut_details_frame,
                                                 textvariable=device_under_test_instrument_id_number_2,
                                                 font=('arial', 12))
        new_additional_serial_number.grid(row=2, column=1)
        new_additional_serial_number.config(width=15)

        # Ask for Customer Asset ID Number of Instrument being Calibrated
        lbl_customer_idn_new = ttk.Label(additional_dut_details_frame, text="Customer Instrument ID Number:",
                                         font=('arial', 12))
        lbl_customer_idn_new.grid(row=3, pady=5, padx=5)
        device_customer_identification_number = ttk.Entry(additional_dut_details_frame,
                                                          textvariable=customer_device_under_test_instrument_number_2,
                                                          font=('arial', 12))
        device_customer_identification_number.grid(row=3, column=1)
        device_customer_identification_number.config(width=15)

        # Ask for Date Code of Instrument being Calibrated
        lbl_new_device_under_test_date_code = ttk.Label(additional_dut_details_frame, text="Date Code:",
                                                        font=('arial', 12))
        lbl_new_device_under_test_date_code.grid(row=4, pady=5)
        device_under_test_date_code_2 = ttk.Entry(additional_dut_details_frame,
                                                  textvariable=device_under_test_date_code_2, font=('arial', 12))
        device_under_test_date_code_2.grid(row=4, column=1)
        device_under_test_date_code_2.config(width=15)

        # User Defined Calibration Date
        lbl_new_calibration_date = ttk.Label(additional_dut_details_frame, text="Calibration Date:", font=('arial', 12))
        lbl_new_calibration_date.grid(pady=5, row=5, padx=5)
        calibration_date_2 = ttk.Entry(additional_dut_details_frame,
                                       textvariable=device_under_test_calibration_date_2, font=12)
        calibration_date_2.grid(row=5, column=1)
        calibration_date_2.config(width=15)

        # User Defined Due Date (Typical 1 Year Spec. Interval)
        lbl_new_due_date = ttk.Label(additional_dut_details_frame, text="Due Date:", font=('arial', 12))
        lbl_new_due_date.grid(pady=3, row=6)
        due_date_2 = ttk.Entry(additional_dut_details_frame, textvariable=device_under_test_due_date_2, font=12)
        due_date_2.grid(row=6, column=1)
        due_date_2.config(width=15)

        # This label is a dummy label that does nothing but help format the
        # window
        dummy = ttk.Label(additional_dut_details_frame)
        dummy.grid(row=1, column=3)
        dummy.config(width=2)

        # ..............................Buttons...................................#

        btn_import_new_certificate_number = ttk.Button(additional_dut_details_frame, text="Import Cert.", width=20,
                                                       command=lambda: self.import_additional_certificate_number())
        btn_import_new_certificate_number.bind("<Return>", lambda event: self.import_additional_certificate_number())
        btn_import_new_certificate_number.grid(pady=5, row=1, column=2)

        btn_import_new_certificate_and_serial_number = ttk.Button(additional_dut_details_frame,
                                                                  text="Import Cert. & Serial", width=20,
                                                                  command=lambda: self.import_additional_certificate_and_serial_number())
        btn_import_new_certificate_and_serial_number.bind("<Return>", lambda event: self.import_additional_certificate_and_serial_number())
        btn_import_new_certificate_and_serial_number.grid(pady=5, row=2, column=2)

        btn_exit_additional_cert_dut = ttk.Button(additional_dut_details_frame, text="Exit", width=20,
                                                  command=lambda: self.exit_additional_calibration())
        btn_exit_additional_cert_dut.bind("<Return>", lambda event: self.exit_additional_calibration())
        btn_exit_additional_cert_dut.grid(pady=5, row=8, column=0, columnspan=2)

        btn_perform_additional_calibration = ttk.Button(additional_dut_details_frame,
                                                        text="Perform Calibration", width=20,
                                                        command=lambda: self.verify_additional_calibration_information_input())
        btn_perform_additional_calibration.bind("<Return>", lambda event: self.verify_additional_calibration_information_input())
        btn_perform_additional_calibration.grid(pady=5, row=8, column=1, columnspan=2)

    # -----------------------------------------------------------------------#

    # Command to Ask User if They Want to Generate Another Certificate of Calibration
    def exit_additional_calibration(self):

        exit_additional_calibration_response = tm.askyesno("Exit Calibration Process?", "If you exit now, you will be \
taken to the main menu. You will have to go through the entire process to generate another certificate. Are you sure \
you want to exit? ")
        if exit_additional_calibration_response is True:
            AdditionalCalCertDUTDetails.destroy()
            LIMSVarConfig.clear_all_certificate_of_calibration_variables()

            from LIMSHomeWindow import AppHomeWindow
            ahw = AppHomeWindow()
            ahw.home_window()

        elif exit_additional_calibration_response is False:
            pass

    # -----------------------------------------------------------------------#

    # Command to Import Additional Certificate Number
    def import_additional_certificate_number(self):
        global new_additional_serial_number

        btn_import_new_certificate_number.config(cursor="watch")
        LIMSVarConfig.certificate_of_calibration_number = ""
        LIMSVarConfig.device_serial_number = ""
        new_additional_certificate_number .config(text=LIMSVarConfig.certificate_of_calibration_number)
        new_additional_serial_number.destroy()

        if LIMSVarConfig.customer_selection_check == int(0):
            LIMSVarConfig.external_customer_sales_order_number_helper = external_customer_sales_order_number_value.get()
        else:
            LIMSVarConfig.external_customer_name = "Dwyer Instruments, Inc."
            LIMSVarConfig.external_customer_sales_order_number_helper = "-"
            LIMSVarConfig.external_customer_rma_number = "-"

        LIMSVarConfig.calibration_date_helper = calibration_date_2.get()
        LIMSVarConfig.calibration_due_date_helper = due_date_2.get()

        try:
            print('Getting current year')
            year = datetime.today().year
            print('Attempting to open database')
            excel_database = open("\\\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls" %(year, year))
            print('Database Opened')
            if excel_database.closed is False:
                excel_database.close()
                new_additional_serial_number = ttk.Entry(additional_dut_details_frame,
                                                         textvariable=device_under_test_instrument_id_number_2,
                                                         font=('arial', 12))
                new_additional_serial_number.grid(row=2, column=1)
                new_additional_serial_number.config(width=15)
                from LIMSCertDBase import AppCertificateDatabase
                acd = AppCertificateDatabase()
                acd.certificate_number_checker()
                if LIMSVarConfig.certificate_of_calibration_number != "":
                    new_additional_certificate_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                else:
                    acd.certificate_number_helper()
                    new_additional_certificate_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                btn_import_new_certificate_number.config(cursor="arrow")
        except IOError as e:
            btn_import_new_certificate_number.config(cursor="arrow")
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------#

    # Command to Import Additional Certificate/Serial Number
    def import_additional_certificate_and_serial_number(self):
        global new_additional_serial_number

        btn_import_new_certificate_and_serial_number.config(cursor="watch")
        LIMSVarConfig.certificate_of_calibration_number = ""
        LIMSVarConfig.device_serial_number = ""
        new_additional_certificate_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
        new_additional_serial_number.destroy()

        if LIMSVarConfig.customer_selection_check == int(0):
            LIMSVarConfig.external_customer_sales_order_number_helper = external_customer_sales_order_number_value.get()
        else:
            LIMSVarConfig.external_customer_name = "Dwyer Instruments, Inc."
            LIMSVarConfig.external_customer_sales_order_number_helper = "-"
            LIMSVarConfig.external_customer_rma_number = "-"

        LIMSVarConfig.calibration_date_helper = calibration_date_2.get()
        LIMSVarConfig.calibration_due_date_helper = due_date_2.get()

        try:
            year = datetime.today().year
            excel_database = open("\\\\\\BDC5\\certdbase\\%s\\%s Certificates of calibration.xls" %(year, year))
            if excel_database.closed is False:
                excel_database.close()
                from LIMSCertDBase import AppCertificateDatabase
                acd = AppCertificateDatabase()
                acd.certificate_serial_number_checker()
                if LIMSVarConfig.certificate_of_calibration_number != "" and LIMSVarConfig.device_serial_number != "":
                    new_additional_certificate_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                    new_additional_serial_number = ttk.Label(additional_dut_details_frame,
                                                             text=LIMSVarConfig.device_serial_number,
                                                             font=('arial', 12))
                    new_additional_serial_number.grid(row=2, column=1)
                    new_additional_serial_number.config(width=15)
                else:
                    acd.certificate_serial_number_helper()
                    new_additional_certificate_number.config(text=LIMSVarConfig.certificate_of_calibration_number)
                    new_additional_serial_number = ttk.Label(additional_dut_details_frame,
                                                             text=LIMSVarConfig.device_serial_number,
                                                             font=('arial', 12))
                    new_additional_serial_number.grid(row=2, column=1)
                    new_additional_serial_number.config(width=15)
                btn_import_new_certificate_and_serial_number.config(cursor="arrow")
        except IOError as e:
            btn_import_new_certificate_and_serial_number.config(cursor="arrow")
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------#

    # Command to Verify Signature/User Fields Have Been Filled Out
    def verify_additional_calibration_information_input(self):

        from LIMSDataImport import AppDataImportModule
        adim = AppDataImportModule()

        if LIMSVarConfig.certificate_of_calibration_number == "" or device_customer_identification_number.get() == "" \
                or device_under_test_date_code_2.get() == "" or calibration_date_2.get() == "" \
                or due_date_2.get() == "":
            tm.showerror("Incomplete Fields", "Please fill out the required information before proceeding.")
        else:
            LIMSVarConfig.device_date_code_helper = device_under_test_date_code_2.get()
            LIMSVarConfig.calibration_date_helper = calibration_date_2.get()
            LIMSVarConfig.calibration_due_date_helper = due_date_2.get()

            if LIMSVarConfig.device_serial_number == "":
                LIMSVarConfig.device_identification_number_helper = new_additional_serial_number.get()
            elif LIMSVarConfig.device_serial_number != "":
                LIMSVarConfig.device_identification_number_helper = LIMSVarConfig.device_serial_number

            if LIMSVarConfig.imported_data_checker == int(1):
                adim.data_import_selection(AdditionalCalCertDUTDetails)
            else:
                self.perform_calibration_on_device_under_test(AdditionalCalCertDUTDetails)
