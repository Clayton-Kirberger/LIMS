"""
LIMSComplaints is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to allow users to document complaints that are originated by the customer or internally.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import os

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppComplaints:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =========================COMPLAINTS DOCUMENTATION========================= #

    # This command is designed to allow the user to log complaints in relation to work performed in the laboratory
    def complaints(self, window):
        window.withdraw()

        global comp_option_sel, complaint_option_selection

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        comp_option_sel = Toplevel()
        comp_option_sel.title("Q.A. - Complaints")
        comp_option_sel.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 305
        height = 135
        screen_width = comp_option_sel.winfo_screenwidth()
        screen_height = comp_option_sel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        comp_option_sel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        comp_option_sel.focus_force()
        comp_option_sel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(comp_option_sel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(comp_option_sel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(comp_option_sel)),
                                           ("Logout", lambda: acc.software_signout(comp_option_sel)),
                                           ("Quit", lambda: acc.software_close(comp_option_sel))])

        # ......................Frame Creation...................... #

        complaint_selection_frame = LabelFrame(comp_option_sel, text="Please select an option from the drop down \
list", relief=SOLID, bd=1, labelanchor="n")
        complaint_selection_frame.grid(row=0, column=0, rowspan=2, padx=10, pady=5)

        # ......................Drop Down Lists..................... #

        if LIMSVarConfig.user_access == 'admin':
            complaint_option_selection = ttk.Combobox(complaint_selection_frame,
                                                      values=[" ", "File Complaint", "View Complaint Log"])
        else:
            complaint_option_selection = ttk.Combobox(complaint_selection_frame,
                                                      values=[" ", "File Complaint"])

        acc.always_active_style(complaint_option_selection)
        complaint_option_selection.configure(state="active", width=19)
        complaint_option_selection.focus()
        complaint_option_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(complaint_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_nc_option_selection = ttk.Button(complaint_selection_frame, text="Open", width=15,
                                                  command=lambda: self.open_complaint_option())
        btn_open_nc_option_selection.bind("<Return>", lambda event: self.open_complaint_option())
        btn_open_nc_option_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_ofi_option_selection = ttk.Button(complaint_selection_frame, text="Back to Main Menu",
                                                      width=20, command=lambda: acc.return_home(comp_option_sel))
        btn_backout_ofi_option_selection.bind("<Return>", lambda event: acc.return_home(comp_option_sel))
        btn_backout_ofi_option_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle complaint selection
    def open_complaint_option(self):
        if complaint_option_selection.get() == "File Complaint":
            self.complaint_identifier_checker()
        elif complaint_option_selection.get() == "View Complaint Log":
            self.complaint_log_viewer()
        else:
            tm.showerror("No Complaint Selection Made", "Please select an action from the drop down provided.")

    # ----------------------------------------------------------------------------- #

    def new_complaint(self, window):
        window.withdraw()

        global new_complaint_info, complaint_location_selection, complaint_severity_selection, \
            complaint_personnel_textbox, complaint_description_textbox, complaint_findings_textbox, \
            btn_submit_complaint

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        acc.obtain_date()

        # ........................Main Window Properties.........................#

        new_complaint_info = Toplevel()
        new_complaint_info.title("New Complaint Creation")
        new_complaint_info.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 622
        height = 465
        screen_width = new_complaint_info.winfo_screenwidth()
        screen_height = new_complaint_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        new_complaint_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        new_complaint_info.focus_force()
        new_complaint_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(new_complaint_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(new_complaint_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(new_complaint_info)),
                                           ("Complaint Selection", lambda: self.complaints(new_complaint_info)),
                                           ("Logout", lambda: acc.software_signout(new_complaint_info)),
                                           ("Quit", lambda: acc.software_close(new_complaint_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.complaints_help())])

        # ..........................Frame Creation...............................#

        new_complaints_frame = LabelFrame(new_complaint_info, text="Complaints", relief=SOLID, bd=1,
                                          labelanchor="n")
        new_complaints_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=5, pady=5)

        # .........................Labels and Entries............................#

        # Date of Creation
        lbl_comp_date_creation = ttk.Label(new_complaints_frame, text="Date of Creation:", font=('arial', 12))
        lbl_comp_date_creation.grid(row=0, pady=5, padx=5)
        complaint_date_value = ttk.Label(new_complaints_frame, text=LIMSVarConfig.date_helper, font=('arial', 12,
                                                                                                     'bold'))
        complaint_date_value.grid(row=0, column=1, pady=5, padx=5)

        # Complaint Identifier
        lbl_comp_identifier = ttk.Label(new_complaints_frame, text="Complaint Number:", font=('arial', 12))
        lbl_comp_identifier.grid(row=0, column=2, pady=5, padx=5)
        complaint_identifier_value = ttk.Label(new_complaints_frame, text=LIMSVarConfig.comp_identifier,
                                               font=('arial', 12, 'bold'))
        complaint_identifier_value.grid(row=0, column=3, pady=5, padx=5)

        # Complaint Location
        lbl_comp_location = ttk.Label(new_complaints_frame, text="Subject/Location:", font=('arial', 12))
        lbl_comp_location.grid(row=1, pady=5, padx=5)
        complaint_location_selection = ttk.Combobox(new_complaints_frame, values=[" ", "Documentation", "Technician",
                                                                                  "P01 - Main Lab", "P01 - Flow Lab",
                                                                                  "Manufacturing",
                                                                                  "Mechanical Room",
                                                                                  "Sustained Pressure Lab",
                                                                                  "SAH Testing Room"])
        acc.always_active_style(complaint_location_selection)
        complaint_location_selection.configure(state="active", width=20)
        complaint_location_selection.focus()
        complaint_location_selection.grid(padx=10, pady=5, row=1, column=1)

        # Complaint Severity
        lbl_comp_severity = ttk.Label(new_complaints_frame, text="Severity:", font=('arial', 12))
        lbl_comp_severity.grid(row=1, column=2, pady=5, padx=5)
        complaint_severity_selection = ttk.Combobox(new_complaints_frame, values=[" ", "Low", "Medium", "High"])
        acc.always_active_style(complaint_severity_selection)
        complaint_severity_selection.configure(state="active", width=20)
        complaint_severity_selection.grid(padx=10, pady=5, row=1, column=3)

        # Personnel Text Box
        comp_personnel_frame = LabelFrame(new_complaints_frame, text="List the personnel the complaint is \
focused on", relief=SOLID, bd=1, labelanchor="n")
        comp_personnel_frame.grid(row=2, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        complaint_personnel_textbox = Text(comp_personnel_frame, wrap=WORD, width=62, height=3)
        complaint_personnel_vscroll = ttk.Scrollbar(comp_personnel_frame, orient='vertical',
                                                    command=complaint_personnel_textbox.yview)
        complaint_personnel_textbox.configure(font=('arial', 11))
        complaint_personnel_textbox['yscroll'] = complaint_personnel_vscroll.set
        complaint_personnel_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        complaint_personnel_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Description Text Box
        comp_description_frame = LabelFrame(new_complaints_frame, text="Provide a detailed description of the \
complaint to be submitted", relief=SOLID, bd=1, labelanchor="n")
        comp_description_frame.grid(row=4, column=0, rowspan=3, columnspan=4, padx=5, pady=5)

        complaint_description_textbox = Text(comp_description_frame, wrap=WORD, width=62, height=4)
        complaint_description_vscroll = ttk.Scrollbar(comp_description_frame, orient='vertical',
                                                      command=complaint_description_textbox.yview)
        complaint_description_textbox.configure(font=('arial', 11))
        complaint_description_textbox['yscroll'] = complaint_description_vscroll.set
        complaint_description_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        complaint_description_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Findings/Evidence Text Box
        comp_findings_frame = LabelFrame(new_complaints_frame, text="Provide objective evidence in support or \
against the complaint being submitted", relief=SOLID, bd=1, labelanchor="n")
        comp_findings_frame.grid(row=7, column=0, rowspan=3, columnspan=4, padx=5, pady=5)

        complaint_findings_textbox = Text(comp_findings_frame, wrap=WORD, width=62, height=4)
        complaint_findings_vscroll = ttk.Scrollbar(comp_findings_frame, orient='vertical',
                                                   command=complaint_findings_textbox.yview)
        complaint_findings_textbox.configure(font=('arial', 11))
        complaint_findings_textbox['yscroll'] = complaint_findings_vscroll.set
        complaint_findings_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        complaint_findings_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_complaint_creation = ttk.Button(new_complaint_info, text="Back", width=20,
                                                    command=lambda: self.complaints(new_complaint_info))
        btn_backout_complaint_creation.bind("<Return>", lambda event: self.complaints(new_complaint_info))
        btn_backout_complaint_creation.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        btn_submit_complaint = ttk.Button(new_complaint_info, text="Submit", width=20,
                                          command=lambda: self.complaint_submission_request())
        btn_submit_complaint.bind("<Return>", lambda event: self.complaint_submission_request())
        btn_submit_complaint.grid(row=4, column=2, columnspan=2, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to allow admin accounts to view existing complaints log
    def complaint_log_viewer(self):

        lims_complaint_log_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Complaints Log.xlsx'
        os.startfile(lims_complaint_log_file)

    # -----------------------------------------------------------------------#

    # Function created to examine existing Complaint log. If no Complaint exists in log, the first Complaint
    # unique identifier will be generated. If a Complaint exists in the log, it will loop until a new
    # Complaint identifier is generated.
    def complaint_identifier_checker(self):

        try:
            complaints_database = open(
                "Required Files\\LIMS Quality Assurance\\LIMS Complaints Log.xlsx",
                "a")
            if complaints_database.closed is False:
                complaints_database.close()
                from datetime import date
                year = date.today().year

                # Initialize location of unique NC identifiers
                complaints_identifiers = []

                excel = win32com.client.dynamic.Dispatch("Excel.Application")
                wkbook = excel.Workbooks.Open(
                    r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Complaints Log.xlsx')
                sheet = wkbook.Sheets("C-L")

                i = 9
                cell = sheet.Cells(i, 2)

                # Start log with new identifier
                if cell.Value is None:
                    x = i
                    LIMSVarConfig.comp_identifier = str(year) + "-DWYC-0001"
                    sheet.Cells(x, 2).Value = LIMSVarConfig.comp_identifier
                    sheet.Cells(x, 14).Value = "Reserved"
                else:
                    # Obtain all identifiers for array
                    for i in range(9, 3000):
                        if cell.Value is not None:
                            complaints_identifiers.append(str(cell).strip().replace(str(year),
                                                                                    "").replace('-DWYC-', ''))
                            i += 1
                            cell = sheet.Cells(i, 2)
                        else:
                            break

                    # Provide new identifier for NC
                    new_identifier = (int(max(complaints_identifiers)) + 1)

                    if len(str(new_identifier)) == 1:
                        new_identifier = str('000') + str(new_identifier)
                    elif len(str(new_identifier)) == 2:
                        new_identifier = str('00') + str(new_identifier)
                    elif len(str(new_identifier)) == 3:
                        new_identifier = str('0') + str(new_identifier)

                    LIMSVarConfig.comp_identifier = str(year) + "-DWYC-" + str(new_identifier)
                    sheet.Cells(i, 2).Value = LIMSVarConfig.comp_identifier
                    sheet.Cells(i, 14).Value = "Reserved"

                wkbook.Close(True)

            self.new_complaint(comp_option_sel)

        except IOError as e:
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------##

    def complaint_submission_request(self):

        LIMSVarConfig.comp_location = complaint_location_selection.get()
        LIMSVarConfig.comp_severity = complaint_severity_selection.get()
        LIMSVarConfig.comp_personnel = complaint_personnel_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.comp_description = complaint_description_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.comp_findings = complaint_findings_textbox.get('1.0', 'end-1c')

        if len(LIMSVarConfig.comp_location) < 4 or len(LIMSVarConfig.comp_severity) < 4 or \
                LIMSVarConfig.comp_personnel == "" or LIMSVarConfig.comp_description == "" or \
                LIMSVarConfig.comp_findings == "":
            tm.showerror("Missing Information", "Some required information is missing. Review all fields \
and make sure everything has been sufficiently filled out.")
        else:
            # try:
            #     ofi_database = open("Required Files\\LIMS Quality Assurance\\LIMS Opportunity for Improvement Log.xlsx", "a")
            #     if ofi_database.closed is False:
            #         ofi_database.close()

            btn_submit_complaint.config(cursor="watch")

            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            wkbook = excel.Workbooks.Open(
                r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Complaints Log.xlsx')
            sheet = wkbook.Sheets("C-L")

            i = 9
            cell = sheet.Cells(i, 2)

            # Start log with new identifier
            for i in range(9, 3000):
                if cell.Value == LIMSVarConfig.comp_identifier:
                    x = i
                    sheet.Cells(x, 3).Value = LIMSVarConfig.certificate_technician_name
                    sheet.Cells(x, 4).Value = LIMSVarConfig.date_helper
                    sheet.Cells(x, 5).Value = LIMSVarConfig.comp_description
                    sheet.Cells(x, 6).Value = LIMSVarConfig.comp_personnel
                    sheet.Cells(x, 7).Value = LIMSVarConfig.comp_location
                    sheet.Cells(x, 8).Value = LIMSVarConfig.comp_severity
                    sheet.Cells(x, 9).Value = LIMSVarConfig.comp_findings
                    sheet.Cells(x, 14).Value = "To Be Reviewed"
                    break
                else:
                    i += 1
                    cell = sheet.Cells(i, 2)

            wkbook.Close(True)

        self.complaint_submission_confirmation(new_complaint_info)

    #             except IOError as e:
    #                 tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
    # moment and try again.")

    # -----------------------------------------------------------------------#

    def complaint_submission_confirmation(self, window):

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        acc.send_email('rmaldonado@dwyermail.com', ['rmaldonado@dwyermail.com', 'RShumaker@dwyermail.com'], [''],
                       'New Complaint Notification', 'A new Complaint (' +
                       LIMSVarConfig.comp_identifier + ') has been submitted for review by LIMS user ' +
                       LIMSVarConfig.certificate_technician_name + '.', 'rmaldonado@dwyermail.com',
                       'riarpivivtxuevtc', 'smtp.gmail.com')

        btn_submit_complaint.config(cursor="arrow")

        tm.showinfo("Complaint Submitted", "Thank you! Your complaint has been submitted and is under review. ")

        acc.return_home(window)
