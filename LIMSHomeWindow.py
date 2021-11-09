"""
LIMSHomeWindow is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to open all other modules that are useful for the laboratory and its daily functions.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import smtplib
import sys

import LIMSVarConfig
from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk


class AppHomeWindow:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =====================HOME WINDOW CHARACTERISTICS============================#

    def home_window(self):
        global Home, btn_coccreation

        from LIMSHelpWindows import AppHelpWindows
        ahw = AppHelpWindows()

        acc = AppCommonCommands()
        Home = Toplevel()
        Home.title("Home")
        Home.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 560
        height = 580
        screen_width = Home.winfo_screenwidth()
        screen_height = Home.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        Home.focus_force()
        Home.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(Home))

        # ........................Label for Home Window...............................#

        lbl_home = Label(Home, text="Laboratory Information Management System - Main Menu",
                         font=('times new roman', 16, 'bold'), width=45)
        lbl_home.pack(pady=1)

        # .........................Menu Bar Creation..................................#

        menubar = MenuBar(Home, width, height, x, y)

        menubar.add_menu("File", commands=[("Logout", lambda: acc.software_signout(Home)),
                                           ("Quit", lambda: acc.software_close(Home))])

        menubar.add_menu("Help", commands=[("About", lambda: self.about_lims()), ("Help", lambda: ahw.home_help())])

        # ........................Frames for Home Window..............................#

        lab_info_frame = LabelFrame(Home, relief=RIDGE, bd=0)
        lab_info_frame.pack(pady=1)

        database_frame = LabelFrame(Home, relief=RIDGE, bd=0)
        database_frame.pack(pady=1)

        quality_frame = LabelFrame(Home, relief=RIDGE, bd=0)
        quality_frame.pack(pady=1)

        # .......................Headers for Home Window..............................#

        lbl_lab = Label(lab_info_frame, text="Laboratory Information", font=('times new roman', 16, 'bold'),
                        relief=SOLID, bd=1, width=30)
        lbl_lab.pack(pady=1)

        lbl_db = Label(database_frame, text="Database Management", font=('times new roman', 16, 'bold'), relief=SOLID,
                       bd=1, width=30)
        lbl_db.pack(pady=1)

        lbl_quality = Label(quality_frame, text="Quality Assurance", font=('times new roman', 16, 'bold'),
                            relief=SOLID, bd=1, width=30)
        lbl_quality.pack(pady=1)

        # ...................Buttons for Access to Other Modules......................#
        # ========================Laboratory Information==============================#

        btn_convcalc = ttk.Button(lab_info_frame, text="Calculators, Converters & Templates", width=45,
                                  command=lambda: self.calc_and_conv(Home))
        btn_convcalc.bind("<Return>", lambda event: self.calc_and_conv(Home))
        btn_convcalc.pack(pady=1)

        btn_uncertainty = ttk.Button(lab_info_frame, text="EMU and TUR Calculator", width=45,
                                     command=lambda: self.emu_and_tur(Home))
        btn_uncertainty.bind("<Return>", lambda event: self.emu_and_tur(Home))
        btn_uncertainty.pack(pady=1)

        btn_procedures = ttk.Button(lab_info_frame, text="SPI's, OP's & Lab Procedures", width=45,
                                    command=lambda: self.spis_and_ops(btn_procedures))
        btn_procedures.bind("<Return>", lambda event: self.spis_and_ops(btn_procedures))
        btn_procedures.pack(pady=1)

        btn_labconditions = ttk.Button(lab_info_frame, text="Environmental Conditions", width=45,
                                       command=lambda: self.environmental_monitor(Home))
        btn_labconditions.bind("<Return>", lambda event: self.environmental_monitor(Home))
        btn_labconditions.pack(pady=1)

        btn_dataacquisition = ttk.Button(lab_info_frame, text="Data Acquisition", width=45,
                                         command=lambda: self.data_acquisition_module(Home))
        btn_dataacquisition.bind("<Return>", lambda event: self.data_acquisition_module(Home))
        btn_dataacquisition.pack(pady=1)

        btn_coccreation = ttk.Button(lab_info_frame, text="Certificate of Calibration Creation", width=45,
                                     command=lambda: self.certificate_of_calibration(Home))
        btn_coccreation.bind("<Return>", lambda event: self.certificate_of_calibration(Home))
        btn_coccreation.pack(pady=1)

        # ==========================Database Information============================== #

        btn_refstandard = ttk.Button(database_frame, text="Calibration Equipment Database", width=45,
                                     command=lambda: self.calibration_equipment_database())
        btn_refstandard.bind("<Return>", lambda event: self.calibration_equipment_database())
        btn_refstandard.pack(pady=1)

        btn_certdbase = ttk.Button(database_frame, text="Certificate Database", width=45,
                                   command=lambda: self.calibration_certificate_database())
        btn_certdbase.bind("<Return>", lambda event: self.calibration_certificate_database())
        btn_certdbase.pack(pady=1)

        btn_rmadbase = ttk.Button(database_frame, text="RMA Database", width=45,
                                  command=lambda: self.returns_database())
        btn_rmadbase.bind("<Return>", lambda event: self.returns_database())
        btn_rmadbase.pack(pady=1)

        # ===========================Quality Information============================== #

        btn_certificateamendment = ttk.Button(quality_frame, text="Certificate of Calibration Amendment", width=45,
                                              command=lambda: self.calibration_certificate_amendment(Home))
        btn_certificateamendment.bind("<Return>", lambda event: self.calibration_certificate_amendment(Home))
        btn_certificateamendment.pack(pady=1)

        btn_complaints = ttk.Button(quality_frame, text="Complaints", width=45,
                                    command=lambda: self.complaints_documentation(Home))
        btn_complaints.bind("<Return>", lambda event: self.complaints_documentation(Home))
        btn_complaints.pack(pady=1)

        btn_correctiveaction = ttk.Button(quality_frame, text="Corrective Actions", width=45,
                                          command=lambda: self.corrective_action_documentation(Home))
        btn_correctiveaction.bind("<Return>", lambda event: self.corrective_action_documentation(Home))
        btn_correctiveaction.pack(pady=1)

        btn_nonconformance = ttk.Button(quality_frame, text="Non-Conformance Documentation", width=45,
                                        command=lambda: self.nonconformance_documentation(Home))
        btn_nonconformance.bind("<Return>", lambda event: self.nonconformance_documentation(Home))
        btn_nonconformance.pack(pady=1)

        btn_oppforimprovement = ttk.Button(quality_frame, text="Opportunity for Improvement", width=45,
                                           command=lambda: self.opp_for_improvement(Home))
        btn_oppforimprovement.bind("<Return>", lambda event: self.opp_for_improvement(Home))
        btn_oppforimprovement.pack(pady=1)

        btn_signout = ttk.Button(Home, text="Sign Out", width=45, command=lambda: acc.software_signout(Home))
        btn_signout.bind("<Return>", lambda event: acc.software_signout(Home))
        btn_signout.pack(pady=10)

    # ----------------------------------------------------------------------------- #

    # Function to provide user information about Dwyer LIMS Software
    def about_lims(self):
        self.__init__()

        tm.showinfo("About LIMS", "Dwyer Instruments, Inc., LIMS (short for Laboratory Information Management System) \
(v. 1.0.2.0) is a graphical user interface (GUI) that allows personnel in the Engineering Laboratory to perform work, \
in a very streamlined, cohesive fashion, as it relates to not only their duties that are outlined in their job \
description, but also work as it relates to a quality system that adheres to the requirements outlined in \
ISO 17025:2017.")

    # -----------------------------------------------------------------------------#

    # Function to initiate Calculators and Converters Window
    def calc_and_conv(self, window):
        self.__init__()

        from LIMSCCT import AppCCT
        new_window = AppCCT()
        new_window.conversions_and_calculators(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate EMU and TUR Uncertainty Analysis Window
    def emu_and_tur(self, window):
        self.__init__()

        from LIMSUnc import AppUncertainty
        new_window = AppUncertainty()
        new_window.uncertainty(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate EMU and TUR Uncertainty Analysis Window
    def spis_and_ops(self, widget):
        self.__init__()
        widget.config(cursor="watch")

        from LIMSSPIsOPs import AppProcedures
        new_window = AppProcedures()
        new_window.procedure_documentation()

        widget.config(cursor="arrow")

    # -----------------------------------------------------------------------------#

    # Function to initiate Environmental Conditions Monitor Window
    def environmental_monitor(self, window):
        self.__init__()

        from LIMSEnvirCond import AppEnvironmentalConditions
        new_window = AppEnvironmentalConditions()
        new_window.lab_conditions(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate Data Acquisition Process Window
    def data_acquisition_module(self, window):
        self.__init__()
        from LIMSDA import AppDataAcquisition
        new_window = AppDataAcquisition()
        new_window.data_acquisition_parameters(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate Equipment Calibration Process Window
    def certificate_of_calibration(self, window):
        self.__init__()

        btn_coccreation.config(cursor="watch")
        from LIMSCertCreation import AppCalibrationModule
        new_window = AppCalibrationModule()
        new_window.customer_calibration_type_selection(window)
        btn_coccreation.config(cursor="arrow")

    # -----------------------------------------------------------------------------#

    # Function to initiate Calibration Equipment Database Window
    def calibration_equipment_database(self):
        self.__init__()

        from LIMSRefStd import AppReferenceStandardDatabase
        new_window = AppReferenceStandardDatabase()
        new_window.calibration_equipment_database_window()

    # -----------------------------------------------------------------------------#

    # Function to initiate Certificate of Calibration Database Window
    def calibration_certificate_database(self):
        self.__init__()

        from LIMSCertDBase import AppCertificateDatabase
        new_window = AppCertificateDatabase()
        new_window.certificate_database()

    # -----------------------------------------------------------------------------#

    # Function to initiate RMA Database Window
    def returns_database(self):
        self.__init__()

        from LIMSRMADBase import AppRMADatabase
        new_window = AppRMADatabase()
        new_window.rma_database()

    # -----------------------------------------------------------------------------#

    # Function to initiate Complaints Window
    def calibration_certificate_amendment(self, window):
        self.__init__()

        from LIMSCertAmend import AppCertificateAmendment
        new_window = AppCertificateAmendment()
        new_window.certificate_of_calibration_amendment(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate Complaints Window
    def complaints_documentation(self, window):
        self.__init__()

        from LIMSComplaints import AppComplaints
        new_window = AppComplaints()
        new_window.complaints(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate NonConformance Window
    def corrective_action_documentation(self, window):
        self.__init__()

        from LIMSCA import AppCorrectiveActions
        new_window = AppCorrectiveActions()
        new_window.corrective_actions(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate NonConformance Window
    def nonconformance_documentation(self, window):
        self.__init__()

        from LIMSNC import AppNonConformance
        new_window = AppNonConformance()
        new_window.nonconforming_work(window)

    # -----------------------------------------------------------------------------#

    # Function to initiate Opportunity for Improvement Window
    def opp_for_improvement(self, window):
        self.__init__()

        from LIMSOFI import AppOpportunityForImprovement
        new_window = AppOpportunityForImprovement()
        new_window.opportunity_for_improvement(window)

    # -----------------------------------------------------------------------------#

    # Function to close the home window. Designed for use in other modules
    def home_window_hide(self):
        self.__init__()

        cmd = AppCommonCommands()
        cmd.hide_window(Home)

    # -----------------------------------------------------------------------------#


######################################################################################

class AppCommonCommands:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # ==============COMMON COMMANDS  & METHODS USED THROUGHOUT PROGRAM================#

    def always_active_style(self, widget):
        self.__init__()
        widget.config(state="active")
        widget.bind("<Leave>", lambda e: "break")

    def all_caps(self, var):
        self.__init__()
        var.set(var.get().upper())

    def obtain_date(self):
        self.__init__()

        from datetime import date

        today = date.today()
        LIMSVarConfig.date_helper = today.strftime("%m/%d/%Y")

    def send_email(self, from_addr, to_addr_list, cc_addr_list, subject, message, login, password,
                   smtpserver):
        self.__init__()

        header = 'From: %s\n' % from_addr
        header += 'To: %s\n' % ','.join(to_addr_list)
        header += 'Cc: %s\n' % ','.join(cc_addr_list)
        header += 'Subject: %s\n\n' % subject
        message = header + message

        server = smtplib.SMTP_SSL(smtpserver, '465')
        server.ehlo()
        server.login(login, password)
        problems = server.sendmail(from_addr, to_addr_list, str(message))
        server.quit()
        return problems

    def destroy_window(self, window):
        self.__init__()
        window.destroy()

    def return_home(self, window):
        self.__init__()
        window.destroy()
        Home.deiconify()
        Home.focus_force()

    def hide_window(self, window):
        self.__init__()
        window.withdraw()

    def on_exit(self, window):
        self.__init__()
        if tm.askyesno("Close Program?", "Do you want to quit the application?") is True:
            window.destroy()
            sys.exit(0)
        else:
            pass

    def software_close(self, window):
        self.__init__()
        window.destroy()
        sys.exit(0)

    def software_signout(self, window):
        self.__init__()
        window.destroy()

        from Dwyer_Engineering_LIMS import LoginWindowRestoration
        cmd = LoginWindowRestoration()
        cmd.root_window_restore()

    def do_nothing(self):
        pass


######################################################################################

class MenuBar:

    # ======================MENU BAR CREATION & STRUCTURE=============================#

    def __init__(self, parent, width, height, x, y):
        self.menubar = Menu(parent)
        self.create_menu(parent, width, height, x, y)

    def create_menu(self, parent, width, height, x, y):
        parent.config(menu=self.menubar)
        parent.geometry("%dx%d+%d+%d" % (width, height, x, y))

    def add_menu(self, menuname, commands):
        menu = Menu(self.menubar, tearoff=0)

        for command in commands:
            menu.add_command(label=command[0], command=command[1])
            if command[0] == "Logout":
                menu.add_separator()
            else:
                pass

        self.menubar.add_cascade(label=menuname, menu=menu)
