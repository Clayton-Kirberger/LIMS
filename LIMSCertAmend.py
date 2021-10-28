"""
LIMSCertAmend is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to allow users to document any changes needing to be made to existing calibration certificates to meet the
customers needs as well as our own internal requirements.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import os

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppCertificateAmendment:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =========================CERTIFICATE OF CALIBRATION AMENDMENT========================= #

    # This command is designed to allow the user to modify existing certificates, create a new revision of previous
    # existing certificate and provide reasoning for modifying document
    def certificate_of_calibration_amendment(self, window):
        self.__init__()
        window.withdraw()

        global cert_amend_option_selection, cert_amend_option_sel

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        cert_amend_option_sel = Toplevel()
        cert_amend_option_sel.title("Q.A. - Certificate Amendment")
        cert_amend_option_sel.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 310
        height = 135
        screen_width = cert_amend_option_sel.winfo_screenwidth()
        screen_height = cert_amend_option_sel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        cert_amend_option_sel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        cert_amend_option_sel.focus_force()
        cert_amend_option_sel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(cert_amend_option_sel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(cert_amend_option_sel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(cert_amend_option_sel)),
                                           ("Logout", lambda: acc.software_signout(cert_amend_option_sel)),
                                           ("Quit", lambda: acc.software_close(cert_amend_option_sel))])

        # ......................Frame Creation...................... #

        cert_amend_selection_frame = LabelFrame(cert_amend_option_sel, text="Please select an option from the drop \
down list", relief=SOLID, bd=1, labelanchor="n")
        cert_amend_selection_frame.grid(row=0, column=0, rowspan=2, padx=10, pady=5)

        # ......................Drop Down Lists..................... #

        if LIMSVarConfig.user_access == 'admin':
            cert_amend_option_selection = ttk.Combobox(cert_amend_selection_frame, values=[" ", "Amend Certificate",
                                                                                           "View Amendment Log"])
        else:
            cert_amend_option_selection = ttk.Combobox(cert_amend_selection_frame, values=[" ", "Amend Certificate"])

        acc.always_active_style(cert_amend_option_selection)
        cert_amend_option_selection.configure(state="active", width=20)
        cert_amend_option_selection.focus()
        cert_amend_option_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(cert_amend_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_cert_amend_option_selection = ttk.Button(cert_amend_selection_frame, text="Open", width=15,
                                                          command=lambda: self.open_certificate_amendment_option())
        btn_open_cert_amend_option_selection.bind("<Return>", lambda event: self.open_certificate_amendment_option())
        btn_open_cert_amend_option_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_cert_amend_option_selection = ttk.Button(cert_amend_selection_frame, text="Back to Main Menu",
                                                             width=20,
                                                             command=lambda: acc.return_home(cert_amend_option_sel))
        btn_backout_cert_amend_option_selection.bind("<Return>", lambda event: acc.return_home(cert_amend_option_sel))
        btn_backout_cert_amend_option_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle certificate amendment selection
    def open_certificate_amendment_option(self):
        if cert_amend_option_selection.get() == "Amend Certificate":
            self.new_certificate_amendment(cert_amend_option_sel)
        elif cert_amend_option_selection.get() == "View Amendment Log":
            self.cert_amend_log_viewer()
        else:
            tm.showerror("No Certificate Amendment Selection Made", "Please select an action from the drop down \
provided.")

    # ----------------------------------------------------------------------------- #

    # Function to handle creating record of amendment and opening amended certificate for changes to be made
    def new_certificate_amendment(self, window):
        self.__init__()
        window.withdraw()

        cal_cert_number = StringVar()

        global cert_original_record_value, cert_amend_requestor_value, cert_amend_rev_number_value, \
            cert_amend_description_textbox, cert_amend_reason_textbox, cert_amend_effect_textbox, \
            btn_submit_cert_amend, new_cert_amend_info

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        acc.obtain_date()

        # ........................Main Window Properties.........................#

        new_cert_amend_info = Toplevel()
        new_cert_amend_info.title("Certificate of Calibration Amendment")
        new_cert_amend_info.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 640
        height = 485
        screen_width = new_cert_amend_info.winfo_screenwidth()
        screen_height = new_cert_amend_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        new_cert_amend_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        new_cert_amend_info.focus_force()
        new_cert_amend_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(new_cert_amend_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(new_cert_amend_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(new_cert_amend_info)),
                                           ("Certificate Amendment Selection",
                                            lambda: self.certificate_of_calibration_amendment(new_cert_amend_info)),
                                           ("Logout", lambda: acc.software_signout(new_cert_amend_info)),
                                           ("Quit", lambda: acc.software_close(new_cert_amend_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.certificate_amendment_help())])

        # ..........................Frame Creation...............................#

        new_certificate_amendment_frame = LabelFrame(new_cert_amend_info, text="Certificate of Calibration Amendment",
                                                     relief=SOLID, bd=1, labelanchor="n")
        new_certificate_amendment_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=5, pady=5)

        # .........................Labels and Entries............................#

        # Date of Amendment
        lbl_cert_amend_date_creation = ttk.Label(new_certificate_amendment_frame, text="Date of Amendment:",
                                                 font=('arial', 12))
        lbl_cert_amend_date_creation.grid(row=0, pady=5, padx=5)
        cert_amend_date_value = ttk.Label(new_certificate_amendment_frame, text=LIMSVarConfig.date_helper,
                                          font=('arial', 12, 'bold'))
        cert_amend_date_value.grid(row=0, column=1, pady=5, padx=5)

        # Original Record
        lbl_cert_original_record = ttk.Label(new_certificate_amendment_frame, text="Original Certificate #:",
                                             font=('arial', 12))
        lbl_cert_original_record.grid(row=0, column=2, pady=5, padx=5)
        cert_original_record_value = ttk.Entry(new_certificate_amendment_frame, textvariable=cal_cert_number)
        cert_original_record_value.bind("<KeyRelease>", lambda event: acc.all_caps(cal_cert_number))
        cert_original_record_value.grid(row=0, column=3, pady=5, padx=5)
        cert_original_record_value.config(width=15, font=('arial', 11))
        cert_original_record_value.focus()

        # Amendment Requestor
        lbl_cert_amend_requestor = ttk.Label(new_certificate_amendment_frame, text="Amendment Requestor:",
                                             font=('arial', 12))
        lbl_cert_amend_requestor.grid(row=1, pady=5, padx=5)
        cert_amend_requestor_value = ttk.Combobox(new_certificate_amendment_frame, values=[" ", "ANAB", "Customer",
                                                                                           "Jason Berry",
                                                                                           "Robert Maldonado",
                                                                                           "Roger Shumaker",
                                                                                           "Steve Burke"])
        acc.always_active_style(cert_amend_requestor_value)
        cert_amend_requestor_value.configure(state="active", width=17)
        cert_amend_requestor_value.grid(padx=5, pady=5, row=1, column=1)

        # Amendment Revision Number
        lbl_amend_rev_number = ttk.Label(new_certificate_amendment_frame, text="Amendment Revision #:",
                                         font=('arial', 12))
        lbl_amend_rev_number.grid(row=1, column=2, pady=5, padx=5)
        cert_amend_rev_number_value = ttk.Combobox(new_certificate_amendment_frame, values=[" ", "1", "2", "3", "4",
                                                                                            "5", "6", "7", "8", "9"])
        acc.always_active_style(cert_amend_rev_number_value)
        cert_amend_rev_number_value.configure(state="active", width=7)
        cert_amend_rev_number_value.grid(row=1, column=3, pady=5, padx=5)

        # Amendment Description
        cert_amend_description_frame = LabelFrame(new_certificate_amendment_frame,
                                                  text="Provide a detailed description of the amendment that needs \
to be made to the original document", relief=SOLID, bd=1, labelanchor="n")
        cert_amend_description_frame.grid(row=2, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        cert_amend_description_textbox = Text(cert_amend_description_frame, wrap=WORD, width=62, height=4)
        cert_amend_description_vscroll = ttk.Scrollbar(cert_amend_description_frame, orient='vertical',
                                                       command=cert_amend_description_textbox.yview)
        cert_amend_description_textbox.configure(font=('arial', 11))
        cert_amend_description_textbox['yscroll'] = cert_amend_description_vscroll.set
        cert_amend_description_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        cert_amend_description_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Reason for Amendment
        cert_amend_reason_frame = LabelFrame(new_certificate_amendment_frame,
                                             text="Provide the reasoning behind the intended amendment \
(Can include requestors reasoning)", relief=SOLID, bd=1, labelanchor="n")
        cert_amend_reason_frame.grid(row=4, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        cert_amend_reason_textbox = Text(cert_amend_reason_frame, wrap=WORD, width=62, height=4)
        cert_amend_reason_vscroll = ttk.Scrollbar(cert_amend_reason_frame, orient='vertical',
                                                  command=cert_amend_reason_textbox.yview)
        cert_amend_reason_textbox.configure(font=('arial', 11))
        cert_amend_reason_textbox['yscroll'] = cert_amend_reason_vscroll.set
        cert_amend_reason_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        cert_amend_reason_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # Potential Effects of Amendment
        cert_amend_effect_frame = LabelFrame(new_certificate_amendment_frame,
                                             text="List the potential effects (if any) of the amendment to \
the original record", relief=SOLID, bd=1, labelanchor="n")
        cert_amend_effect_frame.grid(row=6, column=0, rowspan=2, columnspan=4, padx=5, pady=5)

        cert_amend_effect_textbox = Text(cert_amend_effect_frame, wrap=WORD, width=62, height=4)
        cert_amend_effect_vscroll = ttk.Scrollbar(cert_amend_effect_frame, orient='vertical',
                                                  command=cert_amend_effect_textbox.yview)
        cert_amend_effect_textbox.configure(font=('arial', 11))
        cert_amend_effect_textbox['yscroll'] = cert_amend_effect_vscroll.set
        cert_amend_effect_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        cert_amend_effect_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_cert_amend = ttk.Button(new_cert_amend_info,
                                            text="Back", width=20,
                                            command=lambda: self.certificate_of_calibration_amendment(new_cert_amend_info))
        btn_backout_cert_amend.bind("<Return>",
                                    lambda event: self.certificate_of_calibration_amendment(new_cert_amend_info))
        btn_backout_cert_amend.grid(row=4, column=0, columnspan=2, padx=5, pady=5)

        btn_submit_cert_amend = ttk.Button(new_cert_amend_info, text="Submit", width=20,
                                           command=lambda: self.cert_amend_request())
        btn_submit_cert_amend.bind("<Return>",
                                   lambda event: self.cert_amend_request())
        btn_submit_cert_amend.grid(row=4, column=2, columnspan=2, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to allow admin accounts to view existing Certificate of Calibration Amendment Log
    def cert_amend_log_viewer(self):
        self.__init__()

        lims_cert_amend_log_file = r'\\BDC5\certdbase\2020\2020 Certificates of Calibration - Amendment Log.xlsx'
        os.startfile(lims_cert_amend_log_file)

    # ----------------------------------------------------------------------------- #

    def cert_amend_request(self):
        self.__init__()

        LIMSVarConfig.cert_amend_original_cert = cert_original_record_value.get()
        LIMSVarConfig.cert_amend_request_personnel = cert_amend_requestor_value.get()
        LIMSVarConfig.cert_amend_revision_number = cert_amend_rev_number_value.get()
        LIMSVarConfig.cert_amend_description = cert_amend_description_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.cert_amend_reason = cert_amend_reason_textbox.get('1.0', 'end-1c')
        LIMSVarConfig.cert_amend_effects = cert_amend_effect_textbox.get('1.0', 'end-1c')

        if len(LIMSVarConfig.cert_amend_original_cert) < 12 or LIMSVarConfig.cert_amend_request_personnel == "" or \
                LIMSVarConfig.cert_amend_revision_number == "" or LIMSVarConfig.cert_amend_description == "" or \
                LIMSVarConfig.cert_amend_reason == "" or LIMSVarConfig.cert_amend_effects == "":
            tm.showerror("Missing Information", "Some required information is missing. Review all fields and make \
sure everything has been sufficiently filled out.")
        else:
            try:
                amend_log = open("\\\\BDC5\\certdbase\\2020\\2020 Certificates of Calibration - Amendment Log.xlsx",
                                 "a")
                if amend_log.closed is False:
                    amend_log.close()

                    btn_submit_cert_amend.config(cursor="watch")

                    excel = win32com.client.dynamic.Dispatch("Excel.Application")
                    wkbook = excel.Workbooks.Open(
                        r'\\BDC5\certdbase\2020\2020 Certificates of Calibration - Amendment Log.xlsx')
                    sheet = wkbook.Sheets("Amended Certifications")

                    i = 4
                    cell = sheet.Cells(i, 2)
                    for i in range(4, 3000):
                        if cell.Value is None:
                            x = i
                            sheet.Cells(x, 2).Value = LIMSVarConfig.cert_amend_original_cert
                            sheet.Cells(x, 3).Value = LIMSVarConfig.cert_amend_description
                            sheet.Cells(x, 4).Value = LIMSVarConfig.cert_amend_revision_number
                            sheet.Cells(x, 5).Value = LIMSVarConfig.date_helper
                            sheet.Cells(x, 6).Value = LIMSVarConfig.certificate_technician_name
                            sheet.Cells(x, 7).Value = LIMSVarConfig.cert_amend_request_personnel
                            sheet.Cells(x, 8).Value = LIMSVarConfig.cert_amend_reason
                            sheet.Cells(x, 9).Value = LIMSVarConfig.cert_amend_effects
                            break
                        else:
                            i += 1
                            cell = sheet.Cells(i, 2)

                    wkbook.Close(True)

                    self.cert_amend_submission_confirmation(new_cert_amend_info)

                else:
                    self.__init__()

            except IOError as e:
                tm.showerror("Database Busy!", "Database is currently busy and in use by another user. Please wait a \
moment and try again.")

    # ----------------------------------------------------------------------------- #

    # Function designed to open original certificate, save as amended version, and open amended version
    def open_amended_certificate(self):
        self.__init__()

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        f = '\\\\BDC5\\certdbase\\2020\\' + LIMSVarConfig.cert_amend_original_cert + '.xlsx'
        wb = excel.Workbooks.Open(f)

        ws = wb.Sheets("Single, Plain, AFAL")
        if ws.Range('I10').Value is not None:
            ws.Range('I10').Value = LIMSVarConfig.cert_amend_original_cert + '-' + \
                                    LIMSVarConfig.cert_amend_revision_number
            ws.Range('A45').Value = "This document serves as an amendment to the original, " + \
                                    LIMSVarConfig.cert_amend_original_cert + "."
        else:
            ws = wb.Sheets("Dual, Plain, AFAL")
            if ws.Range('I10').Value is not None:
                ws.Range('I10').Value = LIMSVarConfig.cert_amend_original_cert + '-' + \
                                        LIMSVarConfig.cert_amend_revision_number
                ws.Range('A47').Value = "This document serves as an amendment to the original, " + \
                                        LIMSVarConfig.cert_amend_original_cert + "."
            else:
                ws = wb.Sheets("Transmitter, Plain, AFAL")
                if ws.Range('I10').Value is not None:
                    ws.Range('I10').Value = LIMSVarConfig.cert_amend_original_cert + '-' + \
                                            LIMSVarConfig.cert_amend_revision_number
                    ws.Range('A47').Value = "This document serves as an amendment to the original, " + \
                                            LIMSVarConfig.cert_amend_original_cert + "."
                else:
                    self.__init__()

        new_f = '\\\\BDC5\\certdbase\\2020\\' + LIMSVarConfig.cert_amend_original_cert + '-' + \
                LIMSVarConfig.cert_amend_revision_number + '.xlsx'
        wb.SaveAs(new_f)

        wb.Close(True)

        os.startfile(new_f)

    # ----------------------------------------------------------------------------- #

    def cert_amend_submission_confirmation(self, window):
        self.__init__()

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        acc.send_email('rmaldonado@dwyermail.com', ['rmaldonado@dwyermail.com', 'RShumaker@dwyermail.com'], [''],
                       'New Certificate Amendment Notification', 'A certificate of calibration has been amended (' +
                       LIMSVarConfig.cert_amend_original_cert + '-' + LIMSVarConfig.cert_amend_revision_number +
                       ') and is being prepared for review by LIMS user ' + LIMSVarConfig.certificate_technician_name +
                       '.', 'rmaldonado@dwyermail.com', 'riarpivivtxuevtc', 'smtp.gmail.com')

        btn_submit_cert_amend.config(cursor="arrow")

        tm.showinfo("Certificate Amendment Logged", "Thank you! Your certificate amendment has been identified and \
has been added to the log of amended certificates. Your original certificate (saved as the new amended document) will \
now open for you to make your changes. Be sure to submit your certificate for review prior to official release of \
the document.")

        acc.return_home(window)

        self.open_amended_certificate()
