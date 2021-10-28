"""
LIMSNC is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to allow users to document non-conformances in relation to equipment, work done in the laboratory, or
documentation created and/or provided by Dwyer Instruments, Inc.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import os

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppNonConformance:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =========================NON CONFORMANCE DOCUMENTATION========================= #

    # This command is designed to allow the user to log non-conformances in relation to work performed in the laboratory
    def nonconforming_work(self, window):
        self.__init__()
        window.withdraw()

        global nc_option_sel, nc_option_selection

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        nc_option_sel = Toplevel()
        nc_option_sel.title("Q.A. - N.C.")
        nc_option_sel.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 290
        height = 135
        screen_width = nc_option_sel.winfo_screenwidth()
        screen_height = nc_option_sel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        nc_option_sel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        nc_option_sel.focus_force()
        nc_option_sel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(nc_option_sel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(nc_option_sel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(nc_option_sel)),
                                           ("Logout", lambda: acc.software_signout(nc_option_sel)),
                                           ("Quit", lambda: acc.software_close(nc_option_sel))])

        # ......................Frame Creation...................... #

        nonconformance_selection_frame = LabelFrame(nc_option_sel, text="Please select an option from the drop down \
list", relief=SOLID, bd=1, labelanchor="n")
        nonconformance_selection_frame.grid(row=0, column=0, rowspan=2, padx=10, pady=5)

        # ......................Drop Down Lists..................... #

        if LIMSVarConfig.user_access == 'admin':
            nc_option_selection = ttk.Combobox(nonconformance_selection_frame,
                                               values=[" ", "Generate NC", "View NC Log"])
        else:
            nc_option_selection = ttk.Combobox(nonconformance_selection_frame,
                                               values=[" ", "Generate NC"])

        acc.always_active_style(nc_option_selection)
        nc_option_selection.configure(state="active", width=15)
        nc_option_selection.focus()
        nc_option_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(nonconformance_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_nc_option_selection = ttk.Button(nonconformance_selection_frame, text="Open", width=15,
                                                  command=lambda: self.open_nonconformance_option())
        btn_open_nc_option_selection.bind("<Return>", lambda event: self.open_nonconformance_option())
        btn_open_nc_option_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_nc_option_selection = ttk.Button(nonconformance_selection_frame, text="Back to Main Menu",
                                                      width=20, command=lambda: acc.return_home(nc_option_sel))
        btn_backout_nc_option_selection.bind("<Return>", lambda event: acc.return_home(nc_option_sel))
        btn_backout_nc_option_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle nonconforming work selection
    def open_nonconformance_option(self):
        if nc_option_selection.get() == "Generate NC":
            self.nc_identifier_checker()
        elif nc_option_selection.get() == "View NC Log":
            self.nc_log_viewer()
        else:
            tm.showerror("No N.C. Selection Made", "Please select an action from the drop down provided.")

    # ----------------------------------------------------------------------------- #

    def new_nonconforming_work(self, window):
        self.__init__()
        window.withdraw()

        global new_nc_info, nc_location_selection, nc_risk_selection, nc_personnel_textbox, \
            nc_description_textbox, nc_findings_textbox, btn_submit_nonconformance

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        acc.obtain_date()

        # ........................Main Window Properties.........................#

        new_nc_info = Toplevel()
        new_nc_info.title("New Nonconforming Work Creation")
        new_nc_info.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 585
        height = 480
        screen_width = new_nc_info.winfo_screenwidth()
        screen_height = new_nc_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        new_nc_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        new_nc_info.focus_force()
        new_nc_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(new_nc_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(new_nc_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(new_nc_info)),
                                           ("NC Selection", lambda: self.nonconforming_work(new_nc_info)),
                                           ("Logout", lambda: acc.software_signout(new_nc_info)),
                                           ("Quit", lambda: acc.software_close(new_nc_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.nc_help())])

        # ..........................Frame Creation...............................#

        new_nonconformity_frame = LabelFrame(new_nc_info, text="Nonconforming Work", relief=SOLID, bd=1,
                                             labelanchor="n")
        new_nonconformity_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=5, pady=5)

        # .........................Labels and Entries............................#

        # Date of Creation
        lbl_nc_date_creation = ttk.Label(new_nonconformity_frame, text="Date of Creation:", font=('arial', 12))
        lbl_nc_date_creation.grid(row=0, pady=5, padx=5)
        nc_date_value = ttk.Label(new_nonconformity_frame, text=LIMSVarConfig.date_helper, font=('arial', 12, 'bold'))
        nc_date_value.grid(row=0, column=1, pady=5, padx=5)

        # NC Identifier
        lbl_nc_identifier = ttk.Label(new_nonconformity_frame, text="N.C. Number:", font=('arial', 12))
        lbl_nc_identifier.grid(row=0, column=2, pady=5, padx=5)
        nc_identifier_value = ttk.Label(new_nonconformity_frame, text=LIMSVarConfig.nc_identifier,
                                        font=('arial', 12, 'bold'))
        nc_identifier_value.grid(row=0, column=3, pady=5, padx=5)

        # Location of Incident
        lbl_nc_location = ttk.Label(new_nonconformity_frame, text="Incident Location:", font=('arial', 12))
        lbl_nc_location.grid(row=1, pady=5, padx=5)
        nc_location_selection = ttk.Combobox(new_nonconformity_frame, values=[" ", "P01 - Main Lab",
                                                                              "P01 - Flow Lab",
                                                                              "Manufacturing Floor",
                                                                              "Mechanical Room",
                                                                              "Sustained Pressure Lab",
                                                                              "SAH Testing Room"])
        acc.always_active_style(nc_location_selection)
        nc_location_selection.configure(state="active", width=20)
        nc_location_selection.focus()
        nc_location_selection.grid(padx=10, pady=5, row=1, column=1)

        # Associated Risk Level
        lbl_nc_risk = ttk.Label(new_nonconformity_frame, text="Risk Level:", font=('arial', 12))
        lbl_nc_risk.grid(row=1, column=2, pady=5, padx=5)
        nc_risk_selection = ttk.Combobox(new_nonconformity_frame, values=[" ", "Low Risk / Impact",
                                                                          "Medium Risk / Impact",
                                                                          "High Risk / Impact"])
        acc.always_active_style(nc_risk_selection)
        nc_risk_selection.configure(state="active", width=20)
        nc_risk_selection.grid(padx=10, pady=5, row=1, column=3)

        # Personnel Text box
        nc_personnel_frame = LabelFrame(new_nonconformity_frame,
                                        text="List the personnel involved in the nonconformity", relief=SOLID,
                                        bd=1, labelanchor="n")
        nc_personnel_frame.grid(row=2, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        nc_personnel_textbox = Text(nc_personnel_frame, wrap=WORD, width=62, height=4)
        nc_personnel_vscroll = ttk.Scrollbar(nc_personnel_frame, orient='vertical', command=nc_personnel_textbox.yview)
        nc_personnel_textbox.configure(font=('arial', 11))
        nc_personnel_textbox['yscroll'] = nc_personnel_vscroll.set
        nc_personnel_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        nc_personnel_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Description Text box
        nc_description_frame = LabelFrame(new_nonconformity_frame,
                                          text="Provide a description of the nonconformance incident", relief=SOLID,
                                          bd=1, labelanchor="n")
        nc_description_frame.grid(row=4, column=0, rowspan=3, columnspan=4, padx=5, pady=5)

        nc_description_textbox = Text(nc_description_frame, wrap=WORD, width=62, height=4)
        nc_description_vscroll = ttk.Scrollbar(nc_description_frame, orient='vertical',
                                               command=nc_description_textbox.yview)
        nc_description_textbox.configure(font=('arial', 11))
        nc_description_textbox['yscroll'] = nc_description_vscroll.set
        nc_description_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        nc_description_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Findings/Evidence Text box
        nc_findings_frame = LabelFrame(new_nonconformity_frame,
                                       text="Provide objective evidence in support of the nonconformity claim",
                                       relief=SOLID, bd=1, labelanchor="n")
        nc_findings_frame.grid(row=7, column=0, rowspan=3, columnspan=4, padx=5, pady=5)

        nc_findings_textbox = Text(nc_findings_frame, wrap=WORD, width=62, height=4)
        nc_findings_vscroll = ttk.Scrollbar(nc_findings_frame, orient='vertical',
                                            command=nc_findings_textbox.yview)
        nc_findings_textbox.configure(font=('arial', 11))
        nc_findings_textbox['yscroll'] = nc_findings_vscroll.set
        nc_findings_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        nc_findings_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_nc_creation = ttk.Button(new_nc_info, text="Back", width=20,
                                             command=lambda: self.nonconforming_work(new_nc_info))
        btn_backout_nc_creation.bind("<Return>", lambda event: self.nonconforming_work(new_nc_info))
        btn_backout_nc_creation.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        btn_submit_nonconformance = ttk.Button(new_nc_info, text="Submit", width=20,
                                               command=lambda: self.nc_submission_request())
        btn_submit_nonconformance.bind("<Return>", lambda event: self.nc_submission_request())
        btn_submit_nonconformance.grid(row=4, column=2, columnspan=2, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to allow admin accounts to view existing NC log
    def nc_log_viewer(self):
        self.__init__()

        lims_nc_log_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Nonconformance Log.xlsx'
        os.startfile(lims_nc_log_file)

    # -----------------------------------------------------------------------#

    # Function created to examine existing NC log. If no NC exists in log, the first NC unique identifier will be
    # generated. If an NC exists in the log, it will loop until a new NC identifier is generated.
    def nc_identifier_checker(self):
        self.__init__()

        try:
            nc_database = open(
                "\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\LIMS Quality Assurance\\LIMS Nonconformance Log.xlsx",
                "a")
            if nc_database.closed is False:
                nc_database.close()
                from datetime import date
                year = date.today().year

                # Initialize location of unique NC identifiers
                nc_identifiers = []

                excel = win32com.client.dynamic.Dispatch("Excel.Application")
                wkbook = excel.Workbooks.Open(
                    r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Nonconformance Log.xlsx')
                sheet = wkbook.Sheets("NC-L")

                i = 9
                cell = sheet.Cells(i, 2)

                # Start log with new identifier
                if cell.Value is None:
                    x = i
                    LIMSVarConfig.nc_identifier = str(year) + "-DWYNC-0001"
                    sheet.Cells(x, 2).Value = LIMSVarConfig.nc_identifier
                    sheet.Cells(x, 15).Value = "Reserved"
                else:
                    # Obtain all identifiers for array
                    for i in range(9, 3000):
                        if cell.Value is not None:
                            nc_identifiers.append(str(cell).strip().replace(str(year), "").replace('-DWYNC-', ''))
                            i += 1
                            cell = sheet.Cells(i, 2)
                        else:
                            break

                    # Provide new identifier for NC
                    new_identifier = (int(max(nc_identifiers)) + 1)

                    if len(str(new_identifier)) == 1:
                        new_identifier = str('000') + str(new_identifier)
                    elif len(str(new_identifier)) == 2:
                        new_identifier = str('00') + str(new_identifier)
                    elif len(str(new_identifier)) == 3:
                        new_identifier = str('0') + str(new_identifier)

                    LIMSVarConfig.nc_identifier = str(year) + "-DWYNC-" + str(new_identifier)
                    sheet.Cells(i, 2).Value = LIMSVarConfig.nc_identifier
                    sheet.Cells(i, 15).Value = "Reserved"

                wkbook.Close(True)

            self.new_nonconforming_work(nc_option_sel)

        except IOError as e:
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------##

    def nc_submission_request(self):
        self.__init__()

        LIMSVarConfig.nc_location = nc_location_selection.get()
        LIMSVarConfig.nc_severity = nc_risk_selection.get()
        LIMSVarConfig.nc_personnel = nc_personnel_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.nc_description = nc_description_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.nc_findings = nc_findings_textbox.get('1.0', 'end-1c')

        if len(LIMSVarConfig.nc_location) < 4 or len(LIMSVarConfig.nc_severity) < 4 or \
                LIMSVarConfig.nc_personnel == "" or LIMSVarConfig.nc_description == "" or \
                LIMSVarConfig.nc_findings == "":
            tm.showerror("Missing Information", "Some required information is missing. Review all fields \
and make sure everything has been sufficiently filled out.")
        else:
            # try:
            #     ofi_database = open("\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\LIMS Quality Assurance\\LIMS Opportunity for Improvement Log.xlsx", "a")
            #     if ofi_database.closed is False:
            #         ofi_database.close()

            btn_submit_nonconformance.config(cursor="watch")

            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            wkbook = excel.Workbooks.Open(
                r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Nonconformance Log.xlsx')
            sheet = wkbook.Sheets("NC-L")

            i = 9
            cell = sheet.Cells(i, 2)

            # Start log with new identifier
            for i in range(9, 3000):
                if cell.Value == LIMSVarConfig.nc_identifier:
                    x = i
                    sheet.Cells(x, 3).Value = LIMSVarConfig.date_helper
                    sheet.Cells(x, 4).Value = LIMSVarConfig.certificate_technician_name
                    sheet.Cells(x, 5).Value = LIMSVarConfig.nc_description
                    sheet.Cells(x, 6).Value = LIMSVarConfig.nc_personnel
                    sheet.Cells(x, 7).Value = LIMSVarConfig.nc_location
                    sheet.Cells(x, 8).Value = LIMSVarConfig.nc_severity
                    sheet.Cells(x, 9).Value = LIMSVarConfig.nc_findings
                    sheet.Cells(x, 15).Value = "To Be Reviewed"
                    break
                else:
                    i += 1
                    cell = sheet.Cells(i, 2)

            wkbook.Close(True)

        self.nc_submission_confirmation(new_nc_info)

#             except IOError as e:
#                 tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
# moment and try again.")

    # -----------------------------------------------------------------------#

    def nc_submission_confirmation(self, window):
        self.__init__()

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        acc.send_email('rmaldonado@dwyermail.com', ['rmaldonado@dwyermail.com', 'RShumaker@dwyermail.com'], [''],
                       'New Nonconformance Notification', 'A new Nonconformance (' +
                       LIMSVarConfig.nc_identifier + ') has been submitted for review by LIMS user ' +
                       LIMSVarConfig.certificate_technician_name + '.', 'rmaldonado@dwyermail.com',
                       'riarpivivtxuevtc', 'smtp.gmail.com')

        btn_submit_nonconformance.config(cursor="arrow")

        tm.showinfo("NC Submitted", "Thank you! Your nonconformance has been submitted and is \
under review. ")

        acc.return_home(window)
