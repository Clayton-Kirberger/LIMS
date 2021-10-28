"""
LIMSOFI is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to allow users to document opportunities for improvement that they recognize throughout the laboratory.
This can relate to day to day operations, or specific improvements made to the calibrations performed by the laboratory.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import os

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppOpportunityForImprovement:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =====================OPPORTUNITY FOR IMPROVEMENT DOCUMENTATION======================= #

    # This command is designed to allow the user to log opportunities for improvement
    # in relation to work performed in the laboratory
    def opportunity_for_improvement(self, window):
        self.__init__()
        window.withdraw()

        global OFIOptionSel, ofi_option_selection

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        OFIOptionSel = Toplevel()
        OFIOptionSel.title("Q.A. - O.F.I.")
        OFIOptionSel.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 290
        height = 135
        screen_width = OFIOptionSel.winfo_screenwidth()
        screen_height = OFIOptionSel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        OFIOptionSel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        OFIOptionSel.focus_force()
        OFIOptionSel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(OFIOptionSel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(OFIOptionSel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(OFIOptionSel)),
                                           ("Logout", lambda: acc.software_signout(OFIOptionSel)),
                                           ("Quit", lambda: acc.software_close(OFIOptionSel))])

        # ......................Frame Creation...................... #

        opportunity_selection_frame = LabelFrame(OFIOptionSel, text="Please Select an Option from the Drop Down List",
                                                 relief=SOLID, bd=1, labelanchor="n")
        opportunity_selection_frame.grid(row=0, column=0, rowspan=2, columnspan=2, padx=10, pady=5)

        # .......................Drop Down Lists................................. #

        if LIMSVarConfig.user_access == 'admin':
            ofi_option_selection = ttk.Combobox(opportunity_selection_frame,
                                                values=[" ", "Generate OFI", "View OFI Log"])
        else:
            ofi_option_selection = ttk.Combobox(opportunity_selection_frame,
                                                values=[" ", "Generate OFI"])

        acc.always_active_style(ofi_option_selection)
        ofi_option_selection.configure(state="active", width=15)
        ofi_option_selection.focus()
        ofi_option_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(opportunity_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_ofi_option_selection = ttk.Button(opportunity_selection_frame, text="Open", width=15,
                                                   command=lambda: self.open_opportunity_option())
        btn_open_ofi_option_selection.bind("<Return>", lambda event: self.open_opportunity_option())
        btn_open_ofi_option_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_ofi_option_selection = ttk.Button(opportunity_selection_frame, text="Back to Main Menu", width=20,
                                                      command=lambda: acc.return_home(OFIOptionSel))
        btn_backout_ofi_option_selection.bind("<Return>", lambda event: acc.return_home(OFIOptionSel))
        btn_backout_ofi_option_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle opportunity for improvement type selection
    def open_opportunity_option(self):
        if ofi_option_selection.get() == "Generate OFI":
            self.ofi_identifier_checker()
        elif ofi_option_selection.get() == "View OFI Log":
            self.ofi_log_viewer()
        else:
            tm.showerror("No OFI Selection Made", "Please select an action from the drop down provided.")

    # -----------------------------------------------------------------------#

    def new_opportunity_for_improvement(self, window):
        self.__init__()
        window.withdraw()

        global new_ofi_info, description_textbox, rationale_textbox, impact_textbox, data_textbox, \
            btn_submit_opportunity

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        acc.obtain_date()

        # ........................Main Window Properties.........................#

        new_ofi_info = Toplevel()
        new_ofi_info.title("New Opportunity for Improvement Creation")
        new_ofi_info.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 585
        height = 340
        screen_width = new_ofi_info.winfo_screenwidth()
        screen_height = new_ofi_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        new_ofi_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        new_ofi_info.focus_force()
        new_ofi_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(new_ofi_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(new_ofi_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(new_ofi_info)),
                                           ("OFI Selection", lambda: self.opportunity_for_improvement(new_ofi_info)),
                                           ("Logout", lambda: acc.software_signout(new_ofi_info)),
                                           ("Quit", lambda: acc.software_close(new_ofi_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.ofi_help())])

        # ..........................Frame Creation...............................#

        new_opportunity_frame = LabelFrame(new_ofi_info, text="Opportunity for Improvement", relief=SOLID, bd=1,
                                           labelanchor="n")
        new_opportunity_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=8, pady=5)

        # ........................Notebook Properties............................#

        ofi_nb = ttk.Notebook(new_opportunity_frame)
        ofi_nb.enable_traversal()
        ofi_nb.grid(row=1, column=0, sticky=E+W+N+S, rowspan=2, columnspan=4, pady=5, padx=5)

        # .........................Labels and Entries............................#

        # Date of Creation
        lbl_ofi_date_creation = ttk.Label(new_opportunity_frame, text="Date of Creation:", font=('arial', 12))
        lbl_ofi_date_creation.grid(row=0, pady=5, padx=5)
        ofi_date_value = ttk.Label(new_opportunity_frame, text=LIMSVarConfig.date_helper, font=('arial', 12, 'bold'))
        ofi_date_value.grid(row=0, column=1, pady=5, padx=5)

        # OFI Identifier
        lbl_ofi_identifier = ttk.Label(new_opportunity_frame, text="O.F.I. Number:", font=('arial', 12))
        lbl_ofi_identifier.grid(row=0, column=2, pady=5, padx=5)
        ofi_identifier_value = ttk.Label(new_opportunity_frame, text=LIMSVarConfig.ofi_identifier,
                                         font=('arial', 12, 'bold'))
        ofi_identifier_value.grid(row=0, column=3, pady=5, padx=5)

        # OFI Notebook
        frame = ttk.Frame(ofi_nb)
        frame_1 = ttk.Frame(ofi_nb)
        frame_2 = ttk.Frame(ofi_nb)
        frame_3 = ttk.Frame(ofi_nb)

        ofi_nb.add(frame, text='Description', compound=TOP, padding=2, underline=0)
        ofi_nb.add(frame_1, text='Rationale', compound=TOP, padding=2, underline=0)
        ofi_nb.add(frame_2, text='Impact', compound=TOP, padding=2, underline=0)
        ofi_nb.add(frame_3, text='Data', compound=TOP, padding=2, underline=1)

        # Description Tab
        description_frame = LabelFrame(frame, text="Provide a detailed description for your Opportunity for \
Improvement", relief=SOLID, bd=1, labelanchor="n")
        description_frame.grid(row=0, column=0, padx=5, pady=5)

        description_textbox = Text(description_frame, wrap=WORD, width=62, height=8)
        description_vscroll = ttk.Scrollbar(description_frame, orient='vertical', command=description_textbox.yview)
        description_textbox.configure(font=('arial', 11))
        description_textbox['yscroll'] = description_vscroll.set
        description_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        description_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Rationale Tab
        rationale_frame = LabelFrame(frame_1, text="Provide detailed and valid rationale for your Opportunity for \
Improvement", relief=SOLID, bd=1, labelanchor="n")
        rationale_frame.grid(row=0, column=0, padx=5, pady=5)

        rationale_textbox = Text(rationale_frame, wrap=WORD, width=62, height=8)
        rationale_vscroll = ttk.Scrollbar(rationale_frame, orient='vertical', command=rationale_textbox.yview)
        rationale_textbox.configure(font=('arial', 11))
        rationale_textbox['yscroll'] = rationale_vscroll.set
        rationale_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        rationale_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Impact Tab
        impact_frame = LabelFrame(frame_2, text="Detail the relevance/importance of your Opportunity for Improvement \
and its potential impact", relief=SOLID, bd=1, labelanchor="n")
        impact_frame.grid(row=0, column=0, padx=5, pady=5)

        impact_textbox = Text(impact_frame, wrap=WORD, width=62, height=8)
        impact_vscroll = ttk.Scrollbar(impact_frame, orient='vertical', command=impact_textbox.yview)
        impact_textbox.configure(font=('arial', 11))
        impact_textbox['yscroll'] = impact_vscroll.set
        impact_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        impact_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Data Tab
        data_frame = LabelFrame(frame_3, text="Discuss the data/information in favor of your Opportunity for \
Improvement", relief=SOLID, bd=1, labelanchor="n")
        data_frame.grid(row=0, column=0, padx=5, pady=5)

        data_textbox = Text(data_frame, wrap=WORD, width=62, height=8)
        data_vscroll = ttk.Scrollbar(data_frame, orient='vertical', command=data_textbox.yview)
        data_textbox.configure(font=('arial', 11))
        data_textbox['yscroll'] = data_vscroll.set
        data_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        data_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_ofi_creation = ttk.Button(new_ofi_info, text="Back", width=20,
                                              command=lambda: self.opportunity_for_improvement(new_ofi_info))
        btn_backout_ofi_creation.bind("<Return>", lambda event: self.opportunity_for_improvement(new_ofi_info))
        btn_backout_ofi_creation.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        btn_submit_opportunity = ttk.Button(new_ofi_info, text="Submit", width=20,
                                            command=lambda: self.ofi_submission_request())
        btn_submit_opportunity.bind("<Return>", lambda event: self.ofi_submission_request())
        btn_submit_opportunity.grid(row=4, column=2, columnspan=2, padx=5, pady=5)

    # -----------------------------------------------------------------------#

    # Function created to allow admin accounts to view existing OFI log
    def ofi_log_viewer(self):
        self.__init__()

        lims_ofi_log_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Opportunity for Improvement Log.xlsx'
        os.startfile(lims_ofi_log_file)

    # -----------------------------------------------------------------------#

    # Function created to examine existing OFI log. If no OFI exists in log, the first OFI unique identifier will be
    # generated. If an OFI exists in the log, it will loop until a new OFI identifier is generated.
    def ofi_identifier_checker(self):
        self.__init__()

        try:
            ofi_database = open("\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\LIMS Quality Assurance\\LIMS Opportunity for Improvement Log.xlsx", "a")
            if ofi_database.closed is False:
                ofi_database.close()
                from datetime import date
                year = date.today().year

                # Initialize location of unique OFI identifiers
                ofi_identifiers = []

                excel = win32com.client.dynamic.Dispatch("Excel.Application")
                wkbook = excel.Workbooks.Open(r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Opportunity for Improvement Log.xlsx')
                sheet = wkbook.Sheets("OFI-L")

                i = 9
                cell = sheet.Cells(i, 4)

                # Start log with new identifier
                if cell.Value is None:
                    x = i
                    LIMSVarConfig.ofi_identifier = str(year) + "-DWYOFI-0001"
                    sheet.Cells(x, 4).Value = LIMSVarConfig.ofi_identifier
                    sheet.Cells(x, 11).Value = "Reserved"
                else:
                    # Obtain all identifiers for array
                    for i in range(9, 3000):
                        if cell.Value is not None:
                            ofi_identifiers.append(str(cell).strip().replace(str(year), "").replace('-DWYOFI-', ''))
                            i += 1
                            cell = sheet.Cells(i, 4)
                        else:
                            break

                    # Provide new identifier for OFI
                    new_identifier = (int(max(ofi_identifiers)) + 1)

                    if len(str(new_identifier)) == 1:
                        new_identifier = str('000') + str(new_identifier)
                    elif len(str(new_identifier)) == 2:
                        new_identifier = str('00') + str(new_identifier)
                    elif len(str(new_identifier)) == 3:
                        new_identifier = str('0') + str(new_identifier)

                    LIMSVarConfig.ofi_identifier = str(year) + "-DWYOFI-" + str(new_identifier)
                    sheet.Cells(i, 4).Value = LIMSVarConfig.ofi_identifier
                    sheet.Cells(i, 11).Value = "Reserved"

                wkbook.Close(True)

            self.new_opportunity_for_improvement(OFIOptionSel)

        except IOError as e:
            tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # -----------------------------------------------------------------------##

    def ofi_submission_request(self):
        self.__init__()

        LIMSVarConfig.ofi_description = description_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.ofi_rationale = rationale_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.ofi_impact = impact_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.ofi_data = data_textbox.get('1.0', 'end-1c')

        if LIMSVarConfig.ofi_description == "" or LIMSVarConfig.ofi_rationale == "" or LIMSVarConfig.ofi_impact == "" \
                or LIMSVarConfig.ofi_data == "":
            tm.showerror("Missing Information", "Some required information is missing. Review each of the tabs \
and make sure all fields have been sufficiently filled out.")
        else:
            # try:
            #     ofi_database = open("\\\\BDC5\\Dwyer Engineering LIMS\\Required Files\\LIMS Quality Assurance\\LIMS Opportunity for Improvement Log.xlsx", "a")
            #     if ofi_database.closed is False:
            #         ofi_database.close()

            btn_submit_opportunity.config(cursor="watch")

            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            wkbook = excel.Workbooks.Open(r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS Quality Assurance\LIMS Opportunity for Improvement Log.xlsx')
            sheet = wkbook.Sheets("OFI-L")

            i = 9
            cell = sheet.Cells(i, 4)

            # Start log with new identifier
            for i in range(9, 3000):
                if cell.Value == LIMSVarConfig.ofi_identifier:
                    x = i
                    sheet.Cells(x, 2).Value = LIMSVarConfig.certificate_technician_name
                    sheet.Cells(x, 3).Value = LIMSVarConfig.date_helper
                    sheet.Cells(x, 5).Value = LIMSVarConfig.ofi_description
                    sheet.Cells(x, 6).Value = LIMSVarConfig.ofi_rationale
                    sheet.Cells(x, 7).Value = LIMSVarConfig.ofi_impact
                    sheet.Cells(x, 8).Value = LIMSVarConfig.ofi_data
                    sheet.Cells(x, 11).Value = "To Be Reviewed"
                    break
                else:
                    i += 1
                    cell = sheet.Cells(i, 4)

            wkbook.Close(True)

        self.ofi_submission_confirmation(new_ofi_info)

#             except IOError as e:
#                 tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
# moment and try again.")

    # -----------------------------------------------------------------------#

    def ofi_submission_confirmation(self, window):
        self.__init__()

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        acc.send_email('rmaldonado@dwyermail.com', ['rmaldonado@dwyermail.com', 'RShumaker@dwyermail.com'], [''],
                       'New Opportunity for Improvement Notification', 'A new Opportunity for Improvement (' +
                       LIMSVarConfig.ofi_identifier + ') has been submitted for review by LIMS user ' +
                       LIMSVarConfig.certificate_technician_name + '.', 'rmaldonado@dwyermail.com',
                       'riarpivivtxuevtc', 'smtp.gmail.com')

        btn_submit_opportunity.config(cursor="arrow")

        tm.showinfo("OFI Submitted", "Thank you! Your opportunity for improvement has been submitted and is \
under review. ")

        acc.return_home(window)
