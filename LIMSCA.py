"""
LIMSCA is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to allow users to document corrective actions in relation to non-conforming work documented and/or
effectiveness of any corrective action taken.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import os

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppCorrectiveActions:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =========================CORRECTIVE ACTION DOCUMENTATION========================= #

    # This command is designed to allow the user to log corrective actions in relation to work performed in
    # the laboratory
    def corrective_actions(self, window):
        self.__init__()
        window.withdraw()

        global ca_option_sel, ca_option_selection

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        ca_option_sel = Toplevel()
        ca_option_sel.title("Q.A. - C.A.")
        ca_option_sel.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 290
        height = 135
        screen_width = ca_option_sel.winfo_screenwidth()
        screen_height = ca_option_sel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        ca_option_sel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        ca_option_sel.focus_force()
        ca_option_sel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(ca_option_sel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(ca_option_sel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(ca_option_sel)),
                                           ("Logout", lambda: acc.software_signout(ca_option_sel)),
                                           ("Quit", lambda: acc.software_close(ca_option_sel))])

        # ......................Frame Creation...................... #

        corrective_action_selection_frame = LabelFrame(ca_option_sel, text="Please select an option from the drop down \
list", relief=SOLID, bd=1, labelanchor="n")
        corrective_action_selection_frame.grid(row=0, column=0, rowspan=2, padx=10, pady=5)

        # ......................Drop Down Lists..................... #

        if LIMSVarConfig.user_access == 'admin':
            ca_option_selection = ttk.Combobox(corrective_action_selection_frame,
                                               values=[" ", "Generate CAR", "View CAR Log"])
        else:
            ca_option_selection = ttk.Combobox(corrective_action_selection_frame,
                                               values=[" ", "Generate CAR"])

        acc.always_active_style(ca_option_selection)
        ca_option_selection.configure(state="active", width=15)
        ca_option_selection.focus()
        ca_option_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(corrective_action_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_ca_option_selection = ttk.Button(corrective_action_selection_frame, text="Open", width=15,
                                                  command=lambda: self.open_corrective_action_option())
        btn_open_ca_option_selection.bind("<Return>", lambda event: self.open_corrective_action_option())
        btn_open_ca_option_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_ca_option_selection = ttk.Button(corrective_action_selection_frame, text="Back to Main Menu",
                                                     width=20, command=lambda: acc.return_home(ca_option_sel))
        btn_backout_ca_option_selection.bind("<Return>", lambda event: acc.return_home(ca_option_sel))
        btn_backout_ca_option_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle corrective action selection
    def open_corrective_action_option(self):
        if ca_option_selection.get() == "Generate CAR":
            self.ca_identifier_checker()
        elif ca_option_selection.get() == "View CAR Log":
            self.ca_log_viewer()
        else:
            tm.showerror("No C.A. Selection Made", "Please select an action from the drop down provided.")

    # ----------------------------------------------------------------------------- #

    def new_corrective_action(self, window):
        self.__init__()
        window.withdraw()

        global new_ca_info, ca_investigator_selection, ca_risk_selection, ca_description_textbox, ca_notes_textbox, \
            btn_submit_corrective_action

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        acc.obtain_date()

        # ........................Main Window Properties.........................#

        new_ca_info = Toplevel()
        new_ca_info.title("New Corrective Action Request Creation")
        new_ca_info.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 582
        height = 440
        screen_width = new_ca_info.winfo_screenwidth()
        screen_height = new_ca_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        new_ca_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        new_ca_info.focus_force()
        new_ca_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(new_ca_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(new_ca_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(new_ca_info)),
                                           ("CA Selection", lambda: self.corrective_actions(new_ca_info)),
                                           ("Logout", lambda: acc.software_signout(new_ca_info)),
                                           ("Quit", lambda: acc.software_close(new_ca_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.ca_help())])

        # ..........................Frame Creation...............................#

        new_corrective_action_frame = LabelFrame(new_ca_info, text="Corrective Action Request", relief=SOLID, bd=1,
                                                 labelanchor="n")
        new_corrective_action_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=5, pady=5)

        # .........................Labels and Entries............................#

        # Date of Creation
        lbl_ca_date_creation = ttk.Label(new_corrective_action_frame, text="Date of Creation:", font=('arial', 12))
        lbl_ca_date_creation.grid(row=0, pady=5, padx=5)
        ca_date_value = ttk.Label(new_corrective_action_frame, text=LIMSVarConfig.date_helper, font=('arial', 12,
                                                                                                     'bold'))
        ca_date_value.grid(row=0, column=1, pady=5, padx=5)

        # CA Identifier
        lbl_ca_identifier = ttk.Label(new_corrective_action_frame, text="C.A.R. Number:", font=('arial', 12))
        lbl_ca_identifier.grid(row=0, column=2, pady=5, padx=5)
        ca_identifier_value = ttk.Label(new_corrective_action_frame, text=LIMSVarConfig.ca_identifier,
                                        font=('arial', 12, 'bold'))
        ca_identifier_value.grid(row=0, column=3, pady=5, padx=5)

        # CA Investigator
        lbl_ca_investigator = ttk.Label(new_corrective_action_frame, text="Investigator:", font=('arial', 12))
        lbl_ca_investigator.grid(row=1, pady=5, padx=5)
        ca_investigator_selection = ttk.Combobox(new_corrective_action_frame, values=[" ", "Jason Berry",
                                                                                      "Steve Burke",
                                                                                      "Robert Maldonado",
                                                                                      "Roger Shumaker"])
        acc.always_active_style(ca_investigator_selection)
        ca_investigator_selection.configure(state="active", width=17)
        ca_investigator_selection.focus()
        ca_investigator_selection.grid(padx=10, pady=5, row=1, column=1)

        # Associated Risk Level
        lbl_ca_risk = ttk.Label(new_corrective_action_frame, text="Risk Level:", font=('arial', 12))
        lbl_ca_risk.grid(row=1, column=2, pady=5, padx=5)
        ca_risk_selection = ttk.Combobox(new_corrective_action_frame, values=[" ", "Low Risk / Impact",
                                                                              "Medium Risk / Impact",
                                                                              "High Risk / Impact"])
        acc.always_active_style(ca_risk_selection)
        ca_risk_selection.configure(state="active", width=20)
        ca_risk_selection.grid(padx=10, pady=5, row=1, column=3)

        # Corrective Action Text box
        ca_description_frame = LabelFrame(new_corrective_action_frame,
                                          text="Please detail for your proposed Corrective Action", relief=SOLID,
                                          bd=1, labelanchor="n")
        ca_description_frame.grid(row=2, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        ca_description_textbox = Text(ca_description_frame, wrap=WORD, width=62, height=6)
        ca_description_vscroll = ttk.Scrollbar(ca_description_frame, orient='vertical',
                                               command=ca_description_textbox.yview)
        ca_description_textbox.configure(font=('arial', 11))
        ca_description_textbox['yscroll'] = ca_description_vscroll.set
        ca_description_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        ca_description_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Notes Text box
        ca_notes_frame = LabelFrame(new_corrective_action_frame,
                                    text="Provide any additional information supporting your Corrective Action",
                                    relief=SOLID, bd=1, labelanchor="n")
        ca_notes_frame.grid(row=4, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        ca_notes_textbox = Text(ca_notes_frame, wrap=WORD, width=62, height=6)
        ca_notes_vscroll = ttk.Scrollbar(ca_notes_frame, orient='vertical', command=ca_notes_textbox.yview)
        ca_notes_textbox.configure(font=('arial', 11))
        ca_notes_textbox['yscroll'] = ca_notes_vscroll.set
        ca_notes_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        ca_notes_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_ca_creation = ttk.Button(new_ca_info, text="Back", width=20,
                                             command=lambda: self.corrective_actions(new_ca_info))
        btn_backout_ca_creation.bind("<Return>", lambda event: self.corrective_actions(new_ca_info))
        btn_backout_ca_creation.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        btn_submit_corrective_action = ttk.Button(new_ca_info, text="Submit", width=20,
                                                  command=lambda: self.ca_submission_request())
        btn_submit_corrective_action.bind("<Return>", lambda event: self.ca_submission_request())
        btn_submit_corrective_action.grid(row=4, column=2, columnspan=2, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to allow admin accounts to view existing CA log
    def ca_log_viewer(self):
        self.__init__()

        lims_ca_log_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Corrective Action Request Log.xlsx'
        os.startfile(lims_ca_log_file)

    # -----------------------------------------------------------------------#

    # Function created to examine existing CA log. If no CA exists in log, the first CA unique identifier will be
    # generated. If an CA exists in the log, it will loop until a new CA identifier is generated.
    def ca_identifier_checker(self):
        self.__init__()

        try:
            ca_database = open(
                "Required Files\\LIMS Quality Assurance\\LIMS Corrective Action Request Log.xlsx",
                "a")
            if ca_database.closed is False:
                ca_database.close()
                from datetime import date
                year = date.today().year

                # Initialize location of unique NC identifiers
                ca_identifiers = []

                excel = win32com.client.dynamic.Dispatch("Excel.Application")
                wkbook = excel.Workbooks.Open(
                    r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Corrective Action Request Log.xlsx')
                sheet = wkbook.Sheets("CAR-L")

                i = 4
                cell = sheet.Cells(i, 2)

                # Start log with new identifier
                if cell.Value is None:
                    x = i
                    LIMSVarConfig.ca_identifier = str(year) + "-DWYCAR-0001"
                    sheet.Cells(x, 2).Value = LIMSVarConfig.ca_identifier
                    sheet.Cells(x, 14).Value = "Reserved"
                else:
                    # Obtain all identifiers for array
                    for i in range(4, 3000):
                        if cell.Value is not None:
                            ca_identifiers.append(str(cell).strip().replace(str(year), "").replace('-DWYCAR-', ''))
                            i += 1
                            cell = sheet.Cells(i, 2)
                        else:
                            break

                    # Provide new identifier for NC
                    new_identifier = (int(max(ca_identifiers)) + 1)

                    if len(str(new_identifier)) == 1:
                        new_identifier = str('000') + str(new_identifier)
                    elif len(str(new_identifier)) == 2:
                        new_identifier = str('00') + str(new_identifier)
                    elif len(str(new_identifier)) == 3:
                        new_identifier = str('0') + str(new_identifier)

                    LIMSVarConfig.ca_identifier = str(year) + "-DWYCAR-" + str(new_identifier)
                    sheet.Cells(i, 2).Value = LIMSVarConfig.ca_identifier
                    sheet.Cells(i, 14).Value = "Reserved"

                wkbook.Close(True)

            self.new_corrective_action(ca_option_sel)

        except IOError as e:
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------##

    def ca_submission_request(self):
        self.__init__()

        LIMSVarConfig.ca_investigator = ca_investigator_selection.get()
        LIMSVarConfig.ca_severity = ca_risk_selection.get()
        LIMSVarConfig.ca_description = ca_description_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.ca_notes = ca_notes_textbox.get('1.0', 'end-1c')

        if len(LIMSVarConfig.ca_investigator) < 4 or len(LIMSVarConfig.ca_severity) < 4 or \
                LIMSVarConfig.ca_description == "" or LIMSVarConfig.ca_notes == "":
            tm.showerror("Missing Information", "Some required information is missing. Review all fields \
and make sure everything has been sufficiently filled out.")
        else:
            # try:
            #     ofi_database = open("Required Files\\LIMS Quality Assurance\\LIMS Opportunity for Improvement Log.xlsx", "a")
            #     if ofi_database.closed is False:
            #         ofi_database.close()

            btn_submit_corrective_action.config(cursor="watch")

            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            wkbook = excel.Workbooks.Open(
                r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Corrective Action Request Log.xlsx')
            sheet = wkbook.Sheets("CAR-L")

            i = 4
            cell = sheet.Cells(i, 2)

            # Start log with new identifier
            for i in range(4, 3000):
                if cell.Value == LIMSVarConfig.ca_identifier:
                    x = i
                    sheet.Cells(x, 3).Value = LIMSVarConfig.certificate_technician_name
                    sheet.Cells(x, 4).Value = LIMSVarConfig.date_helper
                    sheet.Cells(x, 5).Value = LIMSVarConfig.ca_investigator
                    sheet.Cells(x, 8).Value = LIMSVarConfig.ca_description
                    sheet.Cells(x, 12).Value = LIMSVarConfig.ca_notes
                    sheet.Cells(x, 13).Value = LIMSVarConfig.ca_severity
                    sheet.Cells(x, 14).Value = "To Be Reviewed"
                    break
                else:
                    i += 1
                    cell = sheet.Cells(i, 2)

            wkbook.Close(True)

        self.ca_submission_confirmation(new_ca_info)

    #             except IOError as e:
    #                 tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
    # moment and try again.")

    # -----------------------------------------------------------------------#

    def ca_submission_confirmation(self, window):
        self.__init__()

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        acc.send_email('rmaldonado@dwyermail.com', ['rmaldonado@dwyermail.com', 'RShumaker@dwyermail.com'], [''],
                       'New Corrective Action Notification', 'A new Corrective Action (' +
                       LIMSVarConfig.ca_identifier + ') has been submitted for review by LIMS user ' +
                       LIMSVarConfig.certificate_technician_name + '. Corrective Action has been assigned to ' +
                       LIMSVarConfig.ca_investigator + ' for initial investigation.', 'rmaldonado@dwyermail.com',
                       'riarpivivtxuevtc', 'smtp.gmail.com')

        btn_submit_corrective_action.config(cursor="arrow")

        tm.showinfo("CA Submitted", "Thank you! Your corrective action has been submitted and is \
under review. ")

        acc.return_home(window)

