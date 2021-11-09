"""
LIMSDA is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of
this module is to allow users to perform data acquisition and log results without generating a certificate of
calibration.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import LIMSVarConfig
import tkinter.messagebox as tm
from tkinter import *
from tkinter import ttk


class AppDataAcquisition:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =========================DATA ACQUISITION MODULE========================= #

    # This command is designed to allow the user to perform data acquisition without generating
    # a calibration certificate.
    def data_acquisition_parameters(self, window):
        self.__init__()
        window.withdraw()

        LIMSVarConfig.data_acquisition_dut_output_type = " "

        global data_acq_parameters, test_type, dut_output, number_of_models, test_environment

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ..........................Main Window Properties........................ #

        data_acq_parameters = Toplevel()
        data_acq_parameters.title("Data Acquisition - Test Setup")
        data_acq_parameters.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 454
        height = 220
        screen_width = data_acq_parameters.winfo_screenwidth()
        screen_height = data_acq_parameters.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        data_acq_parameters.geometry("%dx%d+%d+%d" % (width, height, x, y))
        data_acq_parameters.focus_force()
        data_acq_parameters.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(data_acq_parameters))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(data_acq_parameters, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(data_acq_parameters)),
                                           ("Logout", lambda: acc.software_signout(data_acq_parameters)),
                                           ("Quit", lambda: acc.software_close(data_acq_parameters))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.data_acq_test_setup_help())])

        # ..............................Frame Creation............................ #

        data_acq_setup_frame = LabelFrame(data_acq_parameters, text="Data Acquisition - Test Setup", relief=SOLID,
                                          bd=1, labelanchor="n")
        data_acq_setup_frame.grid(row=0, column=0, rowspan=5, columnspan=2, padx=8, pady=5)

        # .........................Labels, Entries, Dropdowns..................... #

        # //////////////////////Calibration Equipment Setup Frame////////////////// #

        # Ask for Test Type (i.e. Manual, Semi-Auto, Automated)
        lbl_test_type = ttk.Label(data_acq_setup_frame, text="Test Type:", font=('arial', 12))
        lbl_test_type.grid(row=1, pady=5)
        test_type = ttk.Combobox(data_acq_setup_frame, values=[" ", "Manual", "Automated"])
        acc.always_active_style(test_type)
        test_type.configure(state="active", width=20)
        test_type.focus()
        test_type.grid(row=1, column=1)

        # Update Number of DUTs Dropdown
        btn_update_nod = ttk.Button(data_acq_setup_frame, text="Apply", width=20,
                                    command=lambda: self.update_data_acq_tso())
        btn_update_nod.bind("<Return>", lambda event: self.update_data_acq_tso())
        btn_update_nod.grid(row=1, column=2, padx=5, pady=5)

        # Ask for Output Type of DUTs
        lbl_dut_output_type = ttk.Label(data_acq_setup_frame, text="DUT Output Type:", font=('arial', 12))
        lbl_dut_output_type.grid(row=2, pady=5)
        dut_output = ttk.Label(data_acq_setup_frame, text=LIMSVarConfig.data_acquisition_dut_output_type,
                               font=('arial', 12))
        dut_output.grid(row=2, column=1)

        # Ask for Number of Models
        lbl_number_of_models = ttk.Label(data_acq_setup_frame, text="Number of Models:", font=('arial', 12))
        lbl_number_of_models.grid(row=3, padx=5, pady=5)
        number_of_models = ttk.Combobox(data_acq_setup_frame)
        acc.always_active_style(number_of_models)
        number_of_models.config(state="active", width=20)
        number_of_models.grid(row=3, column=1)

        # Ask for Test Environment/Location
        lbl_test_environment = ttk.Label(data_acq_setup_frame, text="Location:", font=('arial', 12), anchor="n")
        lbl_test_environment.config(width=len("Number of DUT"))
        lbl_test_environment.grid(row=4, padx=5, pady=5)
        test_environment = ttk.Combobox(data_acq_setup_frame, values=[" ", "P01 - Flow Lab", "P01 - Main Lab",
                                                                      "Chamber Room", "Manufacturing Floor",
                                                                      "Mechanical Room", "Sustained Pressure Lab",
                                                                      "SAH Testing Room"])
        acc.always_active_style(test_environment)
        test_environment.config(state="active", width=20)
        test_environment.grid(row=4, column=1, padx=2)

        # Dummy label that does nothing but help format the frame
        dummy = ttk.Label(data_acq_setup_frame)
        dummy.grid(row=4, column=2)
        dummy.config(width=2)

        # ................................Buttons....................................#

        btn_backout_data_acq_param = ttk.Button(data_acq_parameters, text="Back", width=15,
                                                command=lambda: acc.return_home(data_acq_parameters))
        btn_backout_data_acq_param.bind("<Return>", lambda event: acc.return_home(data_acq_parameters))
        btn_backout_data_acq_param.grid(row=5, column=0, pady=5)

        btn_next_data_acq_process = ttk.Button(data_acq_parameters, text="Next", width=15,
                                               command=lambda: self.data_acq_param_check())
        btn_next_data_acq_process.bind("<Return>", lambda event: self.data_acq_param_check())
        btn_next_data_acq_process.grid(row=5, column=1, pady=5)

    # -----------------------------------------------------------------------------#

    # Function to update number of model combobox
    def update_data_acq_tso(self):
        self.__init__()

        if test_type.get() == "Manual":
            LIMSVarConfig.data_acquisition_dut_output_type = "Analogue"
            dut_output.configure(text=LIMSVarConfig.data_acquisition_dut_output_type)
            dut_output.grid(row=2, column=1)
            number_of_models.configure(values=[" ", "1"], state="active", width=20)
            number_of_models.set(" ")
            number_of_models.grid(row=3, column=1)
        elif test_type.get() == "Automated":
            LIMSVarConfig.data_acquisition_dut_output_type = "Digital/Transmitter"
            dut_output.configure(text=LIMSVarConfig.data_acquisition_dut_output_type)
            dut_output.grid(row=2, column=1)
            number_of_models.configure(values=[" ", "1"], state="active", width=20)
            number_of_models.set(" ")
            number_of_models.grid(row=3, column=1)
        else:
            LIMSVarConfig.data_acquisition_dut_output_type = ""
            dut_output.configure(text=LIMSVarConfig.data_acquisition_dut_output_type)
            dut_output.grid(row=2, column=1)
            tm.showerror("No Test Type Selection Made", "Please select a test type from the drop down provided.")

    # -----------------------------------------------------------------------------#

    # Function to verify the proper input of the user
    def data_acq_param_check(self):
        self.__init__()

        if test_type.get() == " " or LIMSVarConfig.data_acquisition_dut_output_type == " " \
                or number_of_models.get() == " " or test_environment.get() == " ":
            tm.showerror("Missing Information", "Please fill out the required fields and try again.")
        else:
            LIMSVarConfig.data_acquisition_test_type = test_type.get()
            LIMSVarConfig.data_acquisition_number_of_models = number_of_models.get()
            LIMSVarConfig.data_acquisition_test_environment = test_environment.get()
            self.data_acq_dut_information(data_acq_parameters)

    # -----------------------------------------------------------------------------#

    # Function designed to allow user to provide a profile for the number of models to be tested
    def data_acq_dut_information(self, window):
        self.__init__()
        window.withdraw()

        global data_acq_dut_info, dut_description_1, dut_model_number_1, dut_date_code_1, mode_of_measure_1, \
            units_of_measure_1

        LIMSVarConfig.data_acquisition_dut_description_1 = ""
        LIMSVarConfig.data_acquisition_dut_model_number_1 = ""
        LIMSVarConfig.data_acquisition_dut_date_code_1 = ""
        LIMSVarConfig.data_acquisition_mode_of_measure_1 = ""
        LIMSVarConfig.data_acquisition_dut_selected_units_1 = ""

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ..........................Main Window Properties........................ #

        data_acq_dut_info = Toplevel()
        data_acq_dut_info.title("Data Acquisition - DUT Model Information")
        data_acq_dut_info.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 545
        height = 320
        screen_width = data_acq_dut_info.winfo_screenwidth()
        screen_height = data_acq_dut_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        data_acq_dut_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        data_acq_dut_info.focus_force()
        data_acq_dut_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(data_acq_dut_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(data_acq_dut_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(data_acq_dut_info)),
                                           ("D.A. - Test Setup",
                                            lambda: self.data_acquisition_parameters(data_acq_dut_info)),
                                           ("Logout", lambda: acc.software_signout(data_acq_dut_info)),
                                           ("Quit", lambda: acc.software_close(data_acq_dut_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.data_acq_dut_info_help())])

        # ..............................Frame Creation............................ #

        data_acq_dut_info_frame = LabelFrame(data_acq_dut_info, text="Data Acquisition - DUT Model Information",
                                             relief=SOLID, bd=1, labelanchor="n")
        data_acq_dut_info_frame.grid(row=0, column=0, rowspan=5, columnspan=4, padx=8, pady=5)

        # ........................Notebook Properties............................#

        s = ttk.Style()
        s.configure('white.TLabel', background="white")

        s = ttk.Style()
        s.configure('gray95.TLabel', background="gray95")

        da_di_nb = ttk.Notebook(data_acq_dut_info_frame)
        da_di_nb.enable_traversal()
        da_di_nb.grid(row=0, column=0, sticky=E+W+N+S, rowspan=5, columnspan=5, pady=5, padx=5)

        # Data Acquisition - DUT Model Information Notebook
        da_frame1 = ttk.Notebook(da_di_nb)
        da_frame1.configure(style='gray95.TLabel')
        # da_frame2 = ttk.Notebook(da_di_nb)
        # da_frame3 = ttk.Notebook(da_di_nb)

        if LIMSVarConfig.data_acquisition_number_of_models == "1":
            da_di_nb.add(da_frame1, text="1st Model", compound=TOP, padding=2, underline=0)
        else:
            self.__init__()
        # elif LIMSVarConfig.data_acquisition_number_of_models == "2":
        #     da_di_nb.add(da_frame1, text="1st Model", compound=TOP, padding=2, underline=0)
        #     da_di_nb.add(da_frame2, text="2nd Model", compound=TOP, padding=2, underline=0)
        # else:
        #     da_di_nb.add(da_frame1, text="1st Model", compound=TOP, padding=2, underline=0)
        #     da_di_nb.add(da_frame2, text="2nd Model", compound=TOP, padding=2, underline=0)
        #     da_di_nb.add(da_frame3, text="3rd Model", compound=TOP, padding=2, underline=0)

        # DUT Description for 1st Model
        lbl_dut_description_1 = ttk.Label(da_frame1, text="Description:", font=('arial', 12))
        lbl_dut_description_1.grid(row=0, padx=10, pady=8)
        dut_description_1 = ttk.Entry(da_frame1, width=30)
        dut_description_1.focus()
        dut_description_1.grid(row=0, column=1, columnspan=3, padx=10, pady=8)

        # Model Number for 1st Model
        lbl_dut_model_number_1 = ttk.Label(da_frame1, text="Model Number:", font=('arial', 12))
        lbl_dut_model_number_1.grid(row=1, padx=10, pady=8)
        dut_model_number_1 = ttk.Entry(da_frame1, width=30)
        dut_model_number_1.grid(row=1, column=1, columnspan=3, padx=10, pady=8)

        # Date Code for 1st Model
        lbl_dut_date_code_1 = ttk.Label(da_frame1, text="Date Code:", font=('arial', 12))
        lbl_dut_date_code_1.grid(row=2, padx=10, pady=8)
        dut_date_code_1 = ttk.Entry(da_frame1, width=30)
        dut_date_code_1.grid(row=2, column=1, columnspan=3, padx=10, pady=8)

        # Mode of Measure for 1st Model
        lbl_mode_of_measure_1 = ttk.Label(da_frame1, text="Mode of Measure:", font=('arial', 12))
        lbl_mode_of_measure_1.grid(row=3, padx=10, pady=8)
        mode_of_measure_1 = ttk.Combobox(da_frame1, values=[" ", "Flow", "Pressure", "Relative Humidity",
                                                            "Temperature", "Velocity"])
        acc.always_active_style(mode_of_measure_1)
        mode_of_measure_1.configure(state="active", width=27)
        mode_of_measure_1.grid(row=3, column=1, columnspan=3, padx=10, pady=8)

        # Apply Measurement Mode for Units for 1st Model
        btn_apply_mode = ttk.Button(da_frame1, text="Load Units", width=20,
                                    command=lambda: self.update_data_acq_units())
        btn_apply_mode.bind("<Return>", lambda event: self.update_data_acq_units())
        btn_apply_mode.grid(row=3, column=4, padx=10, pady=8)

        # Units of Measure for 1st Model
        lbl_units_1 = ttk.Label(da_frame1, text="Units of Measure:", font=('arial', 12))
        lbl_units_1.grid(row=4, padx=10, pady=8)
        units_of_measure_1 = ttk.Combobox(da_frame1, values=LIMSVarConfig.data_acquisition_dut_units_of_measure_1)
        acc.always_active_style(units_of_measure_1)
        units_of_measure_1.configure(state="active", width=27)
        units_of_measure_1.grid(row=4, column=1, columnspan=3, padx=10, pady=8)

        # DUT Description for 2nd Model
        # Model Number for 2nd Model
        # Date Code for 2nd Model
        # Mode of Measure for 2nd Model
        # Units of Measure for 2nd Model

        # DUT Description for 3rd Model
        # Model Number for 3rd Model
        # Date Code for 3rd Model
        # Mode of Measure for 3rd Model
        # Units of Measure for 3rd Model

        # ...............................Buttons..................................#

        btn_backout_dut_info = ttk.Button(data_acq_dut_info, text="Back", width=20,
                                          command=lambda: self.data_acquisition_parameters(data_acq_dut_info))
        btn_backout_dut_info.bind("<Return>", lambda event: self.data_acquisition_parameters(data_acq_dut_info))
        btn_backout_dut_info.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

        btn_process_da_dut = ttk.Button(data_acq_dut_info, text="Next", width=20,
                                        command=lambda: self.process_data_acq_dut_information())
        btn_process_da_dut.bind("<Return>", lambda event: self.process_data_acq_dut_information())
        btn_process_da_dut.grid(row=5, column=2, columnspan=2, padx=5, pady=5)

    # -----------------------------------------------------------------------------#

    # Function created to updated units of measure list for Data Acquisition - DUT Model Information notebook
    def update_data_acq_units(self):
        self.__init__()

        if mode_of_measure_1.get() == "Flow":
            LIMSVarConfig.data_acquisition_dut_units_of_measure_1 = [" ", "CCM", "CFM", "ft/s", "GPH", "l/s", "l/h",
                                                                     "LPM", "CMH", "CMS", "ml/min", "SCFH", "SCFM"]
            units_of_measure_1.configure(values=LIMSVarConfig.data_acquisition_dut_units_of_measure_1,
                                         state="active", width=27)
            units_of_measure_1.set(LIMSVarConfig.data_acquisition_dut_units_of_measure_1[0])
            units_of_measure_1.grid(row=4, column=1, columnspan=3, padx=10, pady=8)
        elif mode_of_measure_1.get() == "Pressure":
            LIMSVarConfig.data_acquisition_dut_units_of_measure_1 = [" ", "atm", "bar", "cmH2O", "ftH2O", "hPa",
                                                                     "inH2O", "inH2O at 60" + u"\u00B0" + "F",
                                                                     "inH2O at 20" + u"\u00B0" + "C", "inHg", "kgcm",
                                                                     "kPa", "mbar", "mmH2O", "mmHg", "MPa", "mTorr",
                                                                     "oz/in", "Pa", "psi", "psia", "psid", "psig",
                                                                     "Torr"]
            units_of_measure_1.configure(values=LIMSVarConfig.data_acquisition_dut_units_of_measure_1,
                                         state="active", width=27)
            units_of_measure_1.set(LIMSVarConfig.data_acquisition_dut_units_of_measure_1[0])
            units_of_measure_1.grid(row=4, column=1, columnspan=3, padx=10, pady=8)
        elif mode_of_measure_1.get() == "Relative Humidity":
            LIMSVarConfig.data_acquisition_dut_units_of_measure_1 = [" ", "%RH"]
            units_of_measure_1.configure(values=LIMSVarConfig.data_acquisition_dut_units_of_measure_1,
                                         state="active", width=27)
            units_of_measure_1.set(LIMSVarConfig.data_acquisition_dut_units_of_measure_1[0])
            units_of_measure_1.grid(row=4, column=1, columnspan=3, padx=10, pady=8)
        elif mode_of_measure_1.get() == "Temperature":
            LIMSVarConfig.data_acquisition_dut_units_of_measure_1 = [" ", u"\u00B0" + "C", u"\u00B0" + "F",
                                                                     u"\u00B0" + "K"]
            units_of_measure_1.configure(values=LIMSVarConfig.data_acquisition_dut_units_of_measure_1,
                                         state="active", width=27)
            units_of_measure_1.set(LIMSVarConfig.data_acquisition_dut_units_of_measure_1[0])
            units_of_measure_1.grid(row=4, column=1, columnspan=3, padx=10, pady=8)
        elif mode_of_measure_1.get() == "Velocity":
            LIMSVarConfig.data_acquisition_dut_units_of_measure_1 = [" ", "m/s", "FPM"]
            units_of_measure_1.configure(values=LIMSVarConfig.data_acquisition_dut_units_of_measure_1,
                                         state="active", width=27)
            units_of_measure_1.set(LIMSVarConfig.data_acquisition_dut_units_of_measure_1[0])
            units_of_measure_1.grid(row=4, column=1, columnspan=3, padx=10, pady=8)
        else:
            tm.showerror("No Selection Made", "There appears to be some information missing. Select a unit of measure \
from the drop down provided and try again.")

    # -----------------------------------------------------------------------------#

    # Function created to process Data Acquisition - DUT Model Information input and advance to equipment selection
    def process_data_acq_dut_information(self):
        self.__init__()

        if dut_description_1.get() == "" or dut_model_number_1.get() == "" or dut_date_code_1.get() == "" or \
                mode_of_measure_1.get() == "" or units_of_measure_1.get() == "":
            tm.showerror("Information Missing", "Please review the entry field and drop down selections and ensure \
they are filled out in their entirety before trying again.")
        else:
            LIMSVarConfig.data_acquisition_dut_description_1 = dut_description_1.get()
            LIMSVarConfig.data_acquisition_dut_model_number_1 = dut_model_number_1.get()
            LIMSVarConfig.data_acquisition_dut_date_code_1 = dut_date_code_1.get()
            LIMSVarConfig.data_acquisition_mode_of_measure_1 = mode_of_measure_1.get()
            LIMSVarConfig.data_acquisition_dut_selected_units_1 = units_of_measure_1.get()
            self.data_acq_mte_selection(data_acq_dut_info)

    # -----------------------------------------------------------------------------#

    # Function to select and configure M&TE for data acquisition by user
    def data_acq_mte_selection(self, window):
        self.__init__()
        window.withdraw()

        global mc6_checkvar, mc2_checkvar, hp34401_checkvar, hp34970a_checkvar, mc_daq_checkvar, pace_5000_checkvar, \
            pace_6000_checkvar, ruska_7250lp_checkvar, ts2500_checkvar, cpc2000_checkvar, cpc3000_checkvar, \
            dpi_515_checkvar, mc_812_checkvar, psw4602_checkvar, tse_300_checkvar, mc6_com_value, mc6_baud_selection, \
            mc6_db_selection, mc6_parity_selection, mc6_stopbit_selection, da_mte_selection

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        da_mte_selection = Toplevel()
        da_mte_selection.title("Data Acquisition - M&TE Selection & Configuration")
        da_mte_selection.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 830
        height = 310
        screen_width = da_mte_selection.winfo_screenwidth()
        screen_height = da_mte_selection.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        da_mte_selection.geometry("%dx%d+%d+%d" % (width, height, x, y))
        da_mte_selection.focus_force()
        da_mte_selection.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(da_mte_selection))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(da_mte_selection, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(da_mte_selection)),
                                           ("D.A. - DUT Model Information",
                                            lambda: self.data_acq_dut_information(da_mte_selection)),
                                           ("Logout", lambda: acc.software_signout(da_mte_selection)),
                                           ("Quit", lambda: acc.software_close(da_mte_selection))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.data_acq_mte_selection_help())])
        menubar.add_menu("Test", commands=[("Test Communications", lambda: self.__init__())])

        # ..........................Frame Creation...............................#

        da_mte_selection_frame = LabelFrame(da_mte_selection, text="Data Acquisition - Measurement & Test Equipment",
                                            relief=SOLID, bd=1, labelanchor="n")
        da_mte_selection_frame.grid(row=0, column=0, rowspan=4, columnspan=6, padx=5, pady=5)

        # ........................Notebook Properties............................#

        mte_nb = ttk.Notebook(da_mte_selection_frame)
        mte_nb.enable_traversal()
        mte_nb.grid(row=0, column=0, sticky=E+W+N+S, rowspan=2, columnspan=6, pady=5, padx=5)

        # ..................Labels, Check buttons and Entries.....................#

        # M&TE Notebook
        reference_frame = ttk.Frame(mte_nb)
        source_ref_frame = ttk.Frame(mte_nb)
        source_frame = ttk.Frame(mte_nb)

        mte_nb.add(reference_frame, text="References", compound=TOP, padding=2, underline=0)
        mte_nb.add(source_ref_frame, text="Sources and References", compound=TOP, padding=2, underline=1)
        mte_nb.add(source_frame, text="Sources", compound=TOP, padding=2, underline=0)

        # Reference Tab
        mte_ref_frame = LabelFrame(reference_frame, text="Select the Applicable Reference Standards for your Data \
Acquisition", relief=SOLID, bd=1, labelanchor="n")
        mte_ref_frame.grid(row=0, column=0, padx=5, pady=5)

        # MC6 / ST-2A Row
        # MC6 / ST-2A Checkbutton
        mc6_checkvar = IntVar()
        mc6_checkvar.set(0)
        mc6_chkbtn = ttk.Checkbutton(mte_ref_frame, variable=mc6_checkvar, text="MC6 / ST-2A", offvalue=0, onvalue=1,
                                     command=lambda: self.process_data_acq_chkbtn())
        mc6_chkbtn.grid(row=0, column=0, padx=5, pady=5)

        # MC6 Com Port Entry
        lbl_mc6_com = ttk.Label(mte_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_mc6_com.grid(row=0, column=1, padx=5, pady=5)
        mc6_com_value = ttk.Entry(mte_ref_frame, state="disabled")
        mc6_com_value.config(width=3)
        mc6_com_value.grid(row=0, column=2, padx=5, pady=5)

        # MC6 Baud Rate Drop Down
        lbl_mc6_baud = ttk.Label(mte_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_mc6_baud.grid(row=0, column=3, padx=5, pady=5)
        mc6_baud_selection = ttk.Combobox(mte_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400", "4800",
                                                                 "7200", "9600", "14400", "19200", "38400", "57600",
                                                                 "115200", "230400", "460800", "921600"])
        acc.always_active_style(mc6_baud_selection)
        mc6_baud_selection.configure(state="disabled", width=8)
        mc6_baud_selection.grid(row=0, column=4, padx=5, pady=5)

        # MC6 Data Bits Drop Down
        lbl_mc6_db = ttk.Label(mte_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_mc6_db.grid(row=0, column=5, padx=5, pady=5)
        mc6_db_selection = ttk.Combobox(mte_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(mc6_db_selection)
        mc6_db_selection.configure(state="disabled", width=3)
        mc6_db_selection.grid(row=0, column=6, padx=5, pady=5)

        # MC6 Parity Drop Down
        lbl_mc6_parity = ttk.Label(mte_ref_frame, text="Parity:", font=('arial', 10))
        lbl_mc6_parity.grid(row=0, column=7, padx=5, pady=5)
        mc6_parity_selection = ttk.Combobox(mte_ref_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(mc6_parity_selection)
        mc6_parity_selection.configure(state="disabled", width=7)
        mc6_parity_selection.grid(row=0, column=8, padx=5, pady=5)

        # MC6 Stop Bits Drop Down
        lbl_mc6_sb = ttk.Label(mte_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_mc6_sb.grid(row=0, column=9, padx=5, pady=5)
        mc6_stopbit_selection = ttk.Combobox(mte_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(mc6_stopbit_selection)
        mc6_stopbit_selection.configure(state="disabled", width=7)
        mc6_stopbit_selection.grid(row=0, column=10, padx=5, pady=5)

        # MC2 Row
        mc2_checkvar = IntVar()
        mc2_checkvar.set(0)
        mc2_chkbtn = ttk.Checkbutton(mte_ref_frame, variable=mc2_checkvar, text="MC2", offvalue=0, onvalue=1)
        mc2_chkbtn.configure(state="disabled")
        mc2_chkbtn.grid(row=1, column=0, padx=5, pady=5)

        # MC2 Com Port Entry
        lbl_mc2_com = ttk.Label(mte_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_mc2_com.grid(row=1, column=1, padx=5, pady=5)
        mc2_com_value = ttk.Entry(mte_ref_frame, state="disabled")
        mc2_com_value.config(width=3)
        mc2_com_value.grid(row=1, column=2, padx=5, pady=5)

        # MC2 Baud Rate Drop Down
        lbl_mc2_baud = ttk.Label(mte_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_mc2_baud.grid(row=1, column=3, padx=5, pady=5)
        mc2_baud_selection = ttk.Combobox(mte_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400", "4800",
                                                                 "7200", "9600", "14400", "19200", "38400", "57600",
                                                                 "115200", "230400", "460800", "921600"])
        acc.always_active_style(mc2_baud_selection)
        mc2_baud_selection.configure(state="disabled", width=8)
        mc2_baud_selection.grid(row=1, column=4, padx=5, pady=5)

        # MC2 Data Bits Drop Down
        lbl_mc2_db = ttk.Label(mte_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_mc2_db.grid(row=1, column=5, padx=5, pady=5)
        mc2_db_selection = ttk.Combobox(mte_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(mc2_db_selection)
        mc2_db_selection.configure(state="disabled", width=3)
        mc2_db_selection.grid(row=1, column=6, padx=5, pady=5)

        # MC2 Parity Drop Down
        lbl_mc2_parity = ttk.Label(mte_ref_frame, text="Parity:", font=('arial', 10))
        lbl_mc2_parity.grid(row=1, column=7, padx=5, pady=5)
        mc2_parity_selection = ttk.Combobox(mte_ref_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(mc2_parity_selection)
        mc2_parity_selection.configure(state="disabled", width=7)
        mc2_parity_selection.grid(row=1, column=8, padx=5, pady=5)

        # MC2 Stop Bits Drop Down
        lbl_mc2_sb = ttk.Label(mte_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_mc2_sb.grid(row=1, column=9, padx=5, pady=5)
        mc2_stopbit_selection = ttk.Combobox(mte_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(mc2_stopbit_selection)
        mc2_stopbit_selection.configure(state="disabled", width=7)
        mc2_stopbit_selection.grid(row=1, column=10, padx=5, pady=5)

        # HP34401 Row
        hp34401_checkvar = IntVar()
        hp34401_checkvar.set(0)
        hp34401_chkbtn = ttk.Checkbutton(mte_ref_frame, variable=hp34401_checkvar, text="HP 34401",
                                         offvalue=0, onvalue=1)
        hp34401_chkbtn.configure(state="disabled")
        hp34401_chkbtn.grid(row=2, column=0, padx=5, pady=5)

        # HP34401 Com Port Entry
        lbl_hp34401_com = ttk.Label(mte_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_hp34401_com.grid(row=2, column=1, padx=5, pady=5)
        hp34401_com_value = ttk.Entry(mte_ref_frame, state="disabled")
        hp34401_com_value.config(width=3)
        hp34401_com_value.grid(row=2, column=2, padx=5, pady=5)

        # HP34401 Baud Rate Drop Down
        lbl_hp34401_baud = ttk.Label(mte_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_hp34401_baud.grid(row=2, column=3, padx=5, pady=5)
        hp34401_baud_selection = ttk.Combobox(mte_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400", "4800",
                                                                     "7200", "9600", "14400", "19200", "38400", "57600",
                                                                     "115200", "230400", "460800", "921600"])
        acc.always_active_style(hp34401_baud_selection)
        hp34401_baud_selection.configure(state="disabled", width=8)
        hp34401_baud_selection.grid(row=2, column=4, padx=5, pady=5)

        # HP34401 Data Bits Drop Down
        lbl_hp34401_db = ttk.Label(mte_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_hp34401_db.grid(row=2, column=5, padx=5, pady=5)
        hp34401_db_selection = ttk.Combobox(mte_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(hp34401_db_selection)
        hp34401_db_selection.configure(state="disabled", width=3)
        hp34401_db_selection.grid(row=2, column=6, padx=5, pady=5)

        # HP34401 Parity Drop Down
        lbl_hp34401_parity = ttk.Label(mte_ref_frame, text="Parity:", font=('arial', 10))
        lbl_hp34401_parity.grid(row=2, column=7, padx=5, pady=5)
        hp34401_parity_selection = ttk.Combobox(mte_ref_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(hp34401_parity_selection)
        hp34401_parity_selection.configure(state="disabled", width=7)
        hp34401_parity_selection.grid(row=2, column=8, padx=5, pady=5)

        # HP34401 Stop Bits Drop Down
        lbl_hp34401_sb = ttk.Label(mte_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_hp34401_sb.grid(row=2, column=9, padx=5, pady=5)
        hp34401_stopbit_selection = ttk.Combobox(mte_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(hp34401_stopbit_selection)
        hp34401_stopbit_selection.configure(state="disabled", width=7)
        hp34401_stopbit_selection.grid(row=2, column=10, padx=5, pady=5)

        # HP 34970A Row
        hp34970a_checkvar = IntVar()
        hp34970a_checkvar.set(0)
        hp34970a_chkbtn = ttk.Checkbutton(mte_ref_frame, variable=hp34970a_checkvar, text="HP 34970A",
                                          offvalue=0, onvalue=1)
        hp34970a_chkbtn.config(state="disabled")
        hp34970a_chkbtn.grid(row=3, column=0, padx=5, pady=5)

        # HP34970a Com Port Entry
        lbl_hp34970a_com = ttk.Label(mte_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_hp34970a_com.grid(row=3, column=1, padx=5, pady=5)
        hp34970a_com_value = ttk.Entry(mte_ref_frame, state="disabled")
        hp34970a_com_value.config(width=3)
        hp34970a_com_value.grid(row=3, column=2, padx=5, pady=5)

        # HP34970a Baud Rate Drop Down
        lbl_hp34970a_baud = ttk.Label(mte_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_hp34970a_baud.grid(row=3, column=3, padx=5, pady=5)
        hp34970a_baud_selection = ttk.Combobox(mte_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400", "4800",
                                                                      "7200", "9600", "14400", "19200", "38400",
                                                                      "57600", "115200", "230400", "460800", "921600"])
        acc.always_active_style(hp34970a_baud_selection)
        hp34970a_baud_selection.configure(state="disabled", width=8)
        hp34970a_baud_selection.grid(row=3, column=4, padx=5, pady=5)

        # HP34970a Data Bits Drop Down
        lbl_hp34970a_db = ttk.Label(mte_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_hp34970a_db.grid(row=3, column=5, padx=5, pady=5)
        hp34970a_db_selection = ttk.Combobox(mte_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(hp34970a_db_selection)
        hp34970a_db_selection.configure(state="disabled", width=3)
        hp34970a_db_selection.grid(row=3, column=6, padx=5, pady=5)

        # HP34970a Parity Drop Down
        lbl_hp34970a_parity = ttk.Label(mte_ref_frame, text="Parity:", font=('arial', 10))
        lbl_hp34970a_parity.grid(row=3, column=7, padx=5, pady=5)
        hp34970a_parity_selection = ttk.Combobox(mte_ref_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(hp34970a_parity_selection)
        hp34970a_parity_selection.configure(state="disabled", width=7)
        hp34970a_parity_selection.grid(row=3, column=8, padx=5, pady=5)

        # HP34970a Stop Bits Drop Down
        lbl_hp34970a_sb = ttk.Label(mte_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_hp34970a_sb.grid(row=3, column=9, padx=5, pady=5)
        hp34970a_stopbit_selection = ttk.Combobox(mte_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(hp34970a_stopbit_selection)
        hp34970a_stopbit_selection.configure(state="disabled", width=7)
        hp34970a_stopbit_selection.grid(row=3, column=10, padx=5, pady=5)

        # MC DAQ Row
        mc_daq_checkvar = IntVar()
        mc_daq_checkvar.set(0)
        mc_daq_chkbtn = ttk.Checkbutton(mte_ref_frame, variable=mc_daq_checkvar, text="M.C. DAQ", offvalue=0, onvalue=1)
        mc_daq_chkbtn.config(state="disabled")
        mc_daq_chkbtn.grid(row=4, column=0, padx=5, pady=5)

        # MC DAQ Com Port Entry
        lbl_mc_daq_com = ttk.Label(mte_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_mc_daq_com.grid(row=4, column=1, padx=5, pady=5)
        mc_daq_com_value = ttk.Entry(mte_ref_frame, state="disabled")
        mc_daq_com_value.config(width=3)
        mc_daq_com_value.grid(row=4, column=2, padx=5, pady=5)

        # MC DAQ Baud Rate Drop Down
        lbl_mc_daq_baud = ttk.Label(mte_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_mc_daq_baud.grid(row=4, column=3, padx=5, pady=5)
        mc_daq_baud_selection = ttk.Combobox(mte_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400", "4800",
                                                                    "7200", "9600", "14400", "19200", "38400",
                                                                    "57600", "115200", "230400", "460800", "921600"])
        acc.always_active_style(mc_daq_baud_selection)
        mc_daq_baud_selection.configure(state="disabled", width=8)
        mc_daq_baud_selection.grid(row=4, column=4, padx=5, pady=5)

        # MC DAQ Data Bits Drop Down
        lbl_mc_daq_db = ttk.Label(mte_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_mc_daq_db.grid(row=4, column=5, padx=5, pady=5)
        mc_daq_db_selection = ttk.Combobox(mte_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(mc_daq_db_selection)
        mc_daq_db_selection.configure(state="disabled", width=3)
        mc_daq_db_selection.grid(row=4, column=6, padx=5, pady=5)

        # MC DAQ Parity Drop Down
        lbl_mc_daq_parity = ttk.Label(mte_ref_frame, text="Parity:", font=('arial', 10))
        lbl_mc_daq_parity.grid(row=4, column=7, padx=5, pady=5)
        mc_daq_parity_selection = ttk.Combobox(mte_ref_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(mc_daq_parity_selection)
        mc_daq_parity_selection.configure(state="disabled", width=7)
        mc_daq_parity_selection.grid(row=4, column=8, padx=5, pady=5)

        # MC DAQ Stop Bits Drop Down
        lbl_mc_daq_sb = ttk.Label(mte_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_mc_daq_sb.grid(row=4, column=9, padx=5, pady=5)
        mc_daq_stopbit_selection = ttk.Combobox(mte_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(mc_daq_stopbit_selection)
        mc_daq_stopbit_selection.configure(state="disabled", width=7)
        mc_daq_stopbit_selection.grid(row=4, column=10, padx=5, pady=5)

        # Source and Reference Tab
        mte_source_ref_frame = LabelFrame(source_ref_frame, text="Select the Applicable Reference Standards for your \
Data Acquisition that are both a Source and Reference", relief=SOLID, bd=1, labelanchor="n")
        mte_source_ref_frame.grid(row=0, column=0, padx=5, pady=5)

        # CPC2000 Row
        cpc2000_checkvar = IntVar()
        cpc2000_checkvar.set(0)
        cpc2000_chkbtn = ttk.Checkbutton(mte_source_ref_frame, variable=cpc2000_checkvar,
                                         text="CPC 2000", offvalue=0, onvalue=1)
        cpc2000_chkbtn.configure(state="disabled")
        cpc2000_chkbtn.grid(row=0, column=0, padx=5, pady=5)

        # CPC2000 Com Port Entry
        lbl_cpc2000_com = ttk.Label(mte_source_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_cpc2000_com.grid(row=0, column=1, padx=5, pady=5)
        cpc2000_com_value = ttk.Entry(mte_source_ref_frame, state="disabled")
        cpc2000_com_value.config(width=3)
        cpc2000_com_value.grid(row=0, column=2, padx=5, pady=5)

        # CPC2000 Baud Rate Drop Down
        lbl_cpc2000_baud = ttk.Label(mte_source_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_cpc2000_baud.grid(row=0, column=3, padx=5, pady=5)
        cpc2000_baud_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                            "4800", "7200", "9600", "14400", "19200",
                                                                            "38400", "57600", "115200", "230400",
                                                                            "460800", "921600"])
        acc.always_active_style(cpc2000_baud_selection)
        cpc2000_baud_selection.configure(state="disabled", width=8)
        cpc2000_baud_selection.grid(row=0, column=4, padx=5, pady=5)

        # CPC2000 Data Bits Drop Down
        lbl_cpc2000_db = ttk.Label(mte_source_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_cpc2000_db.grid(row=0, column=5, padx=5, pady=5)
        cpc2000_db_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(cpc2000_db_selection)
        cpc2000_db_selection.configure(state="disabled", width=3)
        cpc2000_db_selection.grid(row=0, column=6, padx=5, pady=5)

        # CPC2000 Parity Drop Down
        lbl_cpc2000_parity = ttk.Label(mte_source_ref_frame, text="Parity:", font=('arial', 10))
        lbl_cpc2000_parity.grid(row=0, column=7, padx=5, pady=5)
        cpc2000_parity_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(cpc2000_parity_selection)
        cpc2000_parity_selection.configure(state="disabled", width=7)
        cpc2000_parity_selection.grid(row=0, column=8, padx=5, pady=5)

        # CPC2000 Stop Bits Drop Down
        lbl_cpc2000_sb = ttk.Label(mte_source_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_cpc2000_sb.grid(row=0, column=9, padx=5, pady=5)
        cpc2000_stopbit_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(cpc2000_stopbit_selection)
        cpc2000_stopbit_selection.configure(state="disabled", width=7)
        cpc2000_stopbit_selection.grid(row=0, column=10, padx=5, pady=5)

        # Pace 5000 Row
        pace_5000_checkvar = IntVar()
        pace_5000_checkvar.set(0)
        pace_5000_chkbtn = ttk.Checkbutton(mte_source_ref_frame, variable=pace_5000_checkvar, text="Pace 5000",
                                           offvalue=0, onvalue=1)
        pace_5000_chkbtn.configure(state="disabled")
        pace_5000_chkbtn.grid(row=1, column=0, padx=5, pady=5)

        # Pace 5000 Com Port Entry
        lbl_pace_5000_com = ttk.Label(mte_source_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_pace_5000_com.grid(row=1, column=1, padx=5, pady=5)
        pace_5000_com_value = ttk.Entry(mte_source_ref_frame, state="disabled")
        pace_5000_com_value.config(width=3)
        pace_5000_com_value.grid(row=1, column=2, padx=5, pady=5)

        # Pace 5000 Baud Rate Drop Down
        lbl_pace_5000_baud = ttk.Label(mte_source_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_pace_5000_baud.grid(row=1, column=3, padx=5, pady=5)
        pace_5000_baud_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "300", "600", "1200", "1800",
                                                                              "2400", "4800", "7200", "9600", "14400",
                                                                              "19200", "38400", "57600", "115200",
                                                                              "230400", "460800", "921600"])
        acc.always_active_style(pace_5000_baud_selection)
        pace_5000_baud_selection.configure(state="disabled", width=8)
        pace_5000_baud_selection.grid(row=1, column=4, padx=5, pady=5)

        # Pace 5000 Data Bits Drop Down
        lbl_pace_5000_db = ttk.Label(mte_source_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_pace_5000_db.grid(row=1, column=5, padx=5, pady=5)
        pace_5000_db_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(pace_5000_db_selection)
        pace_5000_db_selection.configure(state="disabled", width=3)
        pace_5000_db_selection.grid(row=1, column=6, padx=5, pady=5)

        # Pace 5000 Parity Drop Down
        lbl_pace_5000_parity = ttk.Label(mte_source_ref_frame, text="Parity:", font=('arial', 10))
        lbl_pace_5000_parity.grid(row=1, column=7, padx=5, pady=5)
        pace_5000_parity_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "Even", "Odd", "None", "Mark",
                                                                                "Space"])
        acc.always_active_style(pace_5000_parity_selection)
        pace_5000_parity_selection.configure(state="disabled", width=7)
        pace_5000_parity_selection.grid(row=1, column=8, padx=5, pady=5)

        # Pace 5000 Stop Bits Drop Down
        lbl_pace_5000_sb = ttk.Label(mte_source_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_pace_5000_sb.grid(row=1, column=9, padx=5, pady=5)
        pace_5000_stopbit_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(pace_5000_stopbit_selection)
        pace_5000_stopbit_selection.configure(state="disabled", width=7)
        pace_5000_stopbit_selection.grid(row=1, column=10, padx=5, pady=5)

        # Pace 6000 Row
        pace_6000_checkvar = IntVar()
        pace_6000_checkvar.set(0)
        pace_6000_chkbtn = ttk.Checkbutton(mte_source_ref_frame, variable=pace_6000_checkvar, text="Pace 6000",
                                           offvalue=0, onvalue=1)
        pace_6000_chkbtn.configure(state="disabled")
        pace_6000_chkbtn.grid(row=2, column=0, padx=5, pady=5)

        # Pace 6000 Com Port Entry
        lbl_pace_6000_com = ttk.Label(mte_source_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_pace_6000_com.grid(row=2, column=1, padx=5, pady=5)
        pace_6000_com_value = ttk.Entry(mte_source_ref_frame, state="disabled")
        pace_6000_com_value.config(width=3)
        pace_6000_com_value.grid(row=2, column=2, padx=5, pady=5)

        # Pace 6000 Baud Rate Drop Down
        lbl_pace_6000_baud = ttk.Label(mte_source_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_pace_6000_baud.grid(row=2, column=3, padx=5, pady=5)
        pace_6000_baud_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "300", "600", "1200", "1800",
                                                                              "2400", "4800", "7200", "9600", "14400",
                                                                              "19200", "38400", "57600", "115200",
                                                                              "230400", "460800", "921600"])
        acc.always_active_style(pace_6000_baud_selection)
        pace_6000_baud_selection.configure(state="disabled", width=8)
        pace_6000_baud_selection.grid(row=2, column=4, padx=5, pady=5)

        # Pace 6000 Data Bits Drop Down
        lbl_pace_6000_db = ttk.Label(mte_source_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_pace_6000_db.grid(row=2, column=5, padx=5, pady=5)
        pace_6000_db_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(pace_6000_db_selection)
        pace_6000_db_selection.configure(state="disabled", width=3)
        pace_6000_db_selection.grid(row=2, column=6, padx=5, pady=5)

        # Pace 6000 Parity Drop Down
        lbl_pace_6000_parity = ttk.Label(mte_source_ref_frame, text="Parity:", font=('arial', 10))
        lbl_pace_6000_parity.grid(row=2, column=7, padx=5, pady=5)
        pace_6000_parity_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "Even", "Odd", "None", "Mark",
                                                                                "Space"])
        acc.always_active_style(pace_6000_parity_selection)
        pace_6000_parity_selection.configure(state="disabled", width=7)
        pace_6000_parity_selection.grid(row=2, column=8, padx=5, pady=5)

        # Pace 6000 Stop Bits Drop Down
        lbl_pace_6000_sb = ttk.Label(mte_source_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_pace_6000_sb.grid(row=2, column=9, padx=5, pady=5)
        pace_6000_stopbit_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(pace_6000_stopbit_selection)
        pace_6000_stopbit_selection.configure(state="disabled", width=7)
        pace_6000_stopbit_selection.grid(row=2, column=10, padx=5, pady=5)

        # Ruska 7250LP Row
        ruska_7250lp_checkvar = IntVar()
        ruska_7250lp_checkvar.set(0)
        ruska_7250lp_chkbtn = ttk.Checkbutton(mte_source_ref_frame, variable=ruska_7250lp_checkvar, text="Ruska 7250LP",
                                              offvalue=0, onvalue=1)
        ruska_7250lp_chkbtn.configure(state="disabled")
        ruska_7250lp_chkbtn.grid(row=3, column=0, padx=5, pady=5)

        # Ruska 7250LP Com Port Entry
        lbl_ruska_7250lp_com = ttk.Label(mte_source_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_ruska_7250lp_com.grid(row=3, column=1, padx=5, pady=5)
        ruska_7250lp_com_value = ttk.Entry(mte_source_ref_frame, state="disabled")
        ruska_7250lp_com_value.config(width=3)
        ruska_7250lp_com_value.grid(row=3, column=2, padx=5, pady=5)

        # Ruska 7250LP Baud Rate Drop Down
        lbl_ruska_7250lp_baud = ttk.Label(mte_source_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_ruska_7250lp_baud.grid(row=3, column=3, padx=5, pady=5)
        ruska_7250lp_baud_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "300", "600", "1200", "1800",
                                                                                 "2400", "4800", "7200", "9600",
                                                                                 "14400", "19200", "38400", "57600",
                                                                                 "115200", "230400", "460800",
                                                                                 "921600"])
        acc.always_active_style(ruska_7250lp_baud_selection)
        ruska_7250lp_baud_selection.configure(state="disabled", width=8)
        ruska_7250lp_baud_selection.grid(row=3, column=4, padx=5, pady=5)

        # Ruska 7250LP Data Bits Drop Down
        lbl_ruska_7250lp_db = ttk.Label(mte_source_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_ruska_7250lp_db.grid(row=3, column=5, padx=5, pady=5)
        ruska_7250lp_db_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(ruska_7250lp_db_selection)
        ruska_7250lp_db_selection.configure(state="disabled", width=3)
        ruska_7250lp_db_selection.grid(row=3, column=6, padx=5, pady=5)

        # Ruska 7250LP Parity Drop Down
        lbl_ruska_7250lp_parity = ttk.Label(mte_source_ref_frame, text="Parity:", font=('arial', 10))
        lbl_ruska_7250lp_parity.grid(row=3, column=7, padx=5, pady=5)
        ruska_7250lp_parity_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "Even", "Odd", "None", "Mark",
                                                                                   "Space"])
        acc.always_active_style(ruska_7250lp_parity_selection)
        ruska_7250lp_parity_selection.configure(state="disabled", width=7)
        ruska_7250lp_parity_selection.grid(row=3, column=8, padx=5, pady=5)

        # Ruska 7250LP Stop Bits Drop Down
        lbl_ruska_7250lp_sb = ttk.Label(mte_source_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_ruska_7250lp_sb.grid(row=3, column=9, padx=5, pady=5)
        ruska_7250lp_stopbit_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(ruska_7250lp_stopbit_selection)
        ruska_7250lp_stopbit_selection.configure(state="disabled", width=7)
        ruska_7250lp_stopbit_selection.grid(row=3, column=10, padx=5, pady=5)

        # Thunder Scientific Row
        ts2500_checkvar = IntVar()
        ts2500_checkvar.set(0)
        ts2500_chkbtn = ttk.Checkbutton(mte_source_ref_frame, variable=ts2500_checkvar,
                                        text="T.S. 2500", offvalue=0, onvalue=1)
        ts2500_chkbtn.configure(state="disabled")
        ts2500_chkbtn.grid(row=4, column=0, padx=5, pady=5)

        # Thunder Scientific Com Port Entry
        lbl_ts2500_com = ttk.Label(mte_source_ref_frame, text="Comm. Port:", font=('arial', 10))
        lbl_ts2500_com.grid(row=4, column=1, padx=5, pady=5)
        ts2500_com_value = ttk.Entry(mte_source_ref_frame, state="disabled")
        ts2500_com_value.config(width=3)
        ts2500_com_value.grid(row=4, column=2, padx=5, pady=5)

        # Thunder Scientific Baud Rate Drop Down
        lbl_ts2500_baud = ttk.Label(mte_source_ref_frame, text="Baud Rate:", font=('arial', 10))
        lbl_ts2500_baud.grid(row=4, column=3, padx=5, pady=5)
        ts2500_baud_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                           "4800", "7200", "9600", "14400", "19200",
                                                                           "38400", "57600", "115200", "230400",
                                                                           "460800", "921600"])
        acc.always_active_style(ts2500_baud_selection)
        ts2500_baud_selection.configure(state="disabled", width=8)
        ts2500_baud_selection.grid(row=4, column=4, padx=5, pady=5)

        # Thunder Scientific Data Bits Drop Down
        lbl_ts2500_db = ttk.Label(mte_source_ref_frame, text="Data Bits:", font=('arial', 10))
        lbl_ts2500_db.grid(row=4, column=5, padx=5, pady=5)
        ts2500_db_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "7", "8"])
        acc.always_active_style(ts2500_db_selection)
        ts2500_db_selection.configure(state="disabled", width=3)
        ts2500_db_selection.grid(row=4, column=6, padx=5, pady=5)

        # Thunder Scientific Parity Drop Down
        lbl_ts2500_parity = ttk.Label(mte_source_ref_frame, text="Parity:", font=('arial', 10))
        lbl_ts2500_parity.grid(row=4, column=7, padx=5, pady=5)
        ts2500_parity_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "Even", "Odd", "None", "Mark",
                                                                             "Space"])
        acc.always_active_style(ts2500_parity_selection)
        ts2500_parity_selection.configure(state="disabled", width=7)
        ts2500_parity_selection.grid(row=4, column=8, padx=5, pady=5)

        # Thunder Scientific Stop Bits Drop Down
        lbl_ts2500_sb = ttk.Label(mte_source_ref_frame, text="Stop Bits:", font=('arial', 10))
        lbl_ts2500_sb.grid(row=4, column=9, padx=5, pady=5)
        ts2500_stopbit_selection = ttk.Combobox(mte_source_ref_frame, values=[" ", "1", "2"])
        acc.always_active_style(ts2500_stopbit_selection)
        ts2500_stopbit_selection.configure(state="disabled", width=7)
        ts2500_stopbit_selection.grid(row=4, column=10, padx=5, pady=5)

        # Source Tab
        mte_source_frame = LabelFrame(source_frame, text="Select the Applicable Source Standards for your Data \
Acquisition", relief=SOLID, bd=1, labelanchor="n")
        mte_source_frame.grid(row=0, column=0, padx=5, pady=5)

        # CPC3000 Row
        cpc3000_checkvar = IntVar()
        cpc3000_checkvar.set(0)
        cpc3000_chkbtn = ttk.Checkbutton(mte_source_frame, variable=cpc2000_checkvar,
                                         text="CPC 3000", offvalue=0, onvalue=1)
        cpc3000_chkbtn.configure(state="disabled")
        cpc3000_chkbtn.grid(row=0, column=0, padx=5, pady=5)

        # CPC3000 Com Port Entry
        lbl_cpc3000_com = ttk.Label(mte_source_frame, text="Comm. Port:", font=('arial', 10))
        lbl_cpc3000_com.grid(row=0, column=1, padx=5, pady=5)
        cpc3000_com_value = ttk.Entry(mte_source_frame, state="disabled")
        cpc3000_com_value.config(width=3)
        cpc3000_com_value.grid(row=0, column=2, padx=5, pady=5)

        # CPC3000 Baud Rate Drop Down
        lbl_cpc3000_baud = ttk.Label(mte_source_frame, text="Baud Rate:", font=('arial', 10))
        lbl_cpc3000_baud.grid(row=0, column=3, padx=5, pady=5)
        cpc3000_baud_selection = ttk.Combobox(mte_source_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                        "4800", "7200", "9600", "14400", "19200",
                                                                        "38400", "57600", "115200", "230400",
                                                                        "460800", "921600"])
        acc.always_active_style(cpc3000_baud_selection)
        cpc3000_baud_selection.configure(state="disabled", width=8)
        cpc3000_baud_selection.grid(row=0, column=4, padx=5, pady=5)

        # CPC3000 Data Bits Drop Down
        lbl_cpc3000_db = ttk.Label(mte_source_frame, text="Data Bits:", font=('arial', 10))
        lbl_cpc3000_db.grid(row=0, column=5, padx=5, pady=5)
        cpc3000_db_selection = ttk.Combobox(mte_source_frame, values=[" ", "7", "8"])
        acc.always_active_style(cpc3000_db_selection)
        cpc3000_db_selection.configure(state="disabled", width=3)
        cpc3000_db_selection.grid(row=0, column=6, padx=5, pady=5)

        # CPC3000 Parity Drop Down
        lbl_cpc3000_parity = ttk.Label(mte_source_frame, text="Parity:", font=('arial', 10))
        lbl_cpc3000_parity.grid(row=0, column=7, padx=5, pady=5)
        cpc3000_parity_selection = ttk.Combobox(mte_source_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(cpc3000_parity_selection)
        cpc3000_parity_selection.configure(state="disabled", width=7)
        cpc3000_parity_selection.grid(row=0, column=8, padx=5, pady=5)

        # CPC3000 Stop Bits Drop Down
        lbl_cpc3000_sb = ttk.Label(mte_source_frame, text="Stop Bits:", font=('arial', 10))
        lbl_cpc3000_sb.grid(row=0, column=9, padx=5, pady=5)
        cpc3000_stopbit_selection = ttk.Combobox(mte_source_frame, values=[" ", "1", "2"])
        acc.always_active_style(cpc3000_stopbit_selection)
        cpc3000_stopbit_selection.configure(state="disabled", width=7)
        cpc3000_stopbit_selection.grid(row=0, column=10, padx=5, pady=5)

        # DPI 515 Row
        dpi_515_checkvar = IntVar()
        dpi_515_checkvar.set(0)
        dpi_515_chkbtn = ttk.Checkbutton(mte_source_frame, variable=dpi_515_checkvar,
                                         text="DPI 515", offvalue=0, onvalue=1)
        dpi_515_chkbtn.configure(state="disabled")
        dpi_515_chkbtn.grid(row=1, column=0, padx=5, pady=5)

        # DPI 515 Com Port Entry
        lbl_dpi_515_com = ttk.Label(mte_source_frame, text="Comm. Port:", font=('arial', 10))
        lbl_dpi_515_com.grid(row=1, column=1, padx=5, pady=5)
        dpi_515_com_value = ttk.Entry(mte_source_frame, state="disabled")
        dpi_515_com_value.config(width=3)
        dpi_515_com_value.grid(row=1, column=2, padx=5, pady=5)

        # DPI 515 Baud Rate Drop Down
        lbl_dpi_515_baud = ttk.Label(mte_source_frame, text="Baud Rate:", font=('arial', 10))
        lbl_dpi_515_baud.grid(row=1, column=3, padx=5, pady=5)
        dpi_515_baud_selection = ttk.Combobox(mte_source_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                        "4800", "7200", "9600", "14400", "19200",
                                                                        "38400", "57600", "115200", "230400",
                                                                        "460800", "921600"])
        acc.always_active_style(dpi_515_baud_selection)
        dpi_515_baud_selection.configure(state="disabled", width=8)
        dpi_515_baud_selection.grid(row=1, column=4, padx=5, pady=5)

        # DPI 515 Data Bits Drop Down
        lbl_dpi_515_db = ttk.Label(mte_source_frame, text="Data Bits:", font=('arial', 10))
        lbl_dpi_515_db.grid(row=1, column=5, padx=5, pady=5)
        dpi_515_db_selection = ttk.Combobox(mte_source_frame, values=[" ", "7", "8"])
        acc.always_active_style(dpi_515_db_selection)
        dpi_515_db_selection.configure(state="disabled", width=3)
        dpi_515_db_selection.grid(row=1, column=6, padx=5, pady=5)

        # DPI 515 Parity Drop Down
        lbl_dpi_515_parity = ttk.Label(mte_source_frame, text="Parity:", font=('arial', 10))
        lbl_dpi_515_parity.grid(row=1, column=7, padx=5, pady=5)
        dpi_515_parity_selection = ttk.Combobox(mte_source_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(dpi_515_parity_selection)
        dpi_515_parity_selection.configure(state="disabled", width=7)
        dpi_515_parity_selection.grid(row=1, column=8, padx=5, pady=5)

        # DPI 515 Stop Bits Drop Down
        lbl_dpi_515_sb = ttk.Label(mte_source_frame, text="Stop Bits:", font=('arial', 10))
        lbl_dpi_515_sb.grid(row=1, column=9, padx=5, pady=5)
        dpi_515_stopbit_selection = ttk.Combobox(mte_source_frame, values=[" ", "1", "2"])
        acc.always_active_style(dpi_515_stopbit_selection)
        dpi_515_stopbit_selection.configure(state="disabled", width=7)
        dpi_515_stopbit_selection.grid(row=1, column=10, padx=5, pady=5)

        # Espec MC-812
        mc_812_checkvar = IntVar()
        mc_812_checkvar.set(0)
        mc_812_chkbtn = ttk.Checkbutton(mte_source_frame, variable=mc_812_checkvar, text="Espec MC-812",
                                        offvalue=0, onvalue=1)
        mc_812_chkbtn.configure(state="disabled")
        mc_812_chkbtn.grid(row=2, column=0, padx=5, pady=5)

        # MC-812 Com Port Entry
        lbl_mc_812_com = ttk.Label(mte_source_frame, text="Comm. Port:", font=('arial', 10))
        lbl_mc_812_com.grid(row=2, column=1, padx=5, pady=5)
        mc_812_com_value = ttk.Entry(mte_source_frame, state="disabled")
        mc_812_com_value.config(width=3)
        mc_812_com_value.grid(row=2, column=2, padx=5, pady=5)

        # MC-812 Baud Rate Drop Down
        lbl_mc_812_baud = ttk.Label(mte_source_frame, text="Baud Rate:", font=('arial', 10))
        lbl_mc_812_baud.grid(row=2, column=3, padx=5, pady=5)
        mc_812_baud_selection = ttk.Combobox(mte_source_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                       "4800", "7200", "9600", "14400", "19200",
                                                                       "38400", "57600", "115200", "230400",
                                                                       "460800", "921600"])
        acc.always_active_style(mc_812_baud_selection)
        mc_812_baud_selection.configure(state="disabled", width=8)
        mc_812_baud_selection.grid(row=2, column=4, padx=5, pady=5)

        # MC-812 Data Bits Drop Down
        lbl_mc_812_db = ttk.Label(mte_source_frame, text="Data Bits:", font=('arial', 10))
        lbl_mc_812_db.grid(row=2, column=5, padx=5, pady=5)
        mc_812_db_selection = ttk.Combobox(mte_source_frame, values=[" ", "7", "8"])
        acc.always_active_style(mc_812_db_selection)
        mc_812_db_selection.configure(state="disabled", width=3)
        mc_812_db_selection.grid(row=2, column=6, padx=5, pady=5)

        # MC-812 Parity Drop Down
        lbl_mc_812_parity = ttk.Label(mte_source_frame, text="Parity:", font=('arial', 10))
        lbl_mc_812_parity.grid(row=2, column=7, padx=5, pady=5)
        mc_812_parity_selection = ttk.Combobox(mte_source_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(mc_812_parity_selection)
        mc_812_parity_selection.configure(state="disabled", width=7)
        mc_812_parity_selection.grid(row=2, column=8, padx=5, pady=5)

        # MC-812 Stop Bits Drop Down
        lbl_mc_812_sb = ttk.Label(mte_source_frame, text="Stop Bits:", font=('arial', 10))
        lbl_mc_812_sb.grid(row=2, column=9, padx=5, pady=5)
        mc_812_stopbit_selection = ttk.Combobox(mte_source_frame, values=[" ", "1", "2"])
        acc.always_active_style(mc_812_stopbit_selection)
        mc_812_stopbit_selection.configure(state="disabled", width=7)
        mc_812_stopbit_selection.grid(row=2, column=10, padx=5, pady=5)

        # Espec TSE-12
        # Tektronix PSW4602 Row
        psw4602_checkvar = IntVar()
        psw4602_checkvar.set(0)
        psw4602_chkbtn = ttk.Checkbutton(mte_source_frame, variable=psw4602_checkvar, text="PSW4602", offvalue=0,
                                         onvalue=1)
        psw4602_chkbtn.configure(state="disabled")
        psw4602_chkbtn.grid(row=4, column=0, padx=5, pady=5)

        # PSW4602 Com Port Entry
        lbl_psw4602_com = ttk.Label(mte_source_frame, text="Comm. Port:", font=('arial', 10))
        lbl_psw4602_com.grid(row=4, column=1, padx=5, pady=5)
        psw4602_com_value = ttk.Entry(mte_source_frame, state="disabled")
        psw4602_com_value.config(width=3)
        psw4602_com_value.grid(row=4, column=2, padx=5, pady=5)

        # PSW4602 Baud Rate Drop Down
        lbl_psw4602_baud = ttk.Label(mte_source_frame, text="Baud Rate:", font=('arial', 10))
        lbl_psw4602_baud.grid(row=4, column=3, padx=5, pady=5)
        psw4602_baud_selection = ttk.Combobox(mte_source_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                        "4800", "7200", "9600", "14400", "19200",
                                                                        "38400", "57600", "115200", "230400",
                                                                        "460800", "921600"])
        acc.always_active_style(psw4602_baud_selection)
        psw4602_baud_selection.configure(state="disabled", width=8)
        psw4602_baud_selection.grid(row=4, column=4, padx=5, pady=5)

        # PSW4602 Data Bits Drop Down
        lbl_psw4602_db = ttk.Label(mte_source_frame, text="Data Bits:", font=('arial', 10))
        lbl_psw4602_db.grid(row=4, column=5, padx=5, pady=5)
        psw4602_db_selection = ttk.Combobox(mte_source_frame, values=[" ", "7", "8"])
        acc.always_active_style(psw4602_db_selection)
        psw4602_db_selection.configure(state="disabled", width=3)
        psw4602_db_selection.grid(row=4, column=6, padx=5, pady=5)

        # PSW4602 Parity Drop Down
        lbl_psw4602_parity = ttk.Label(mte_source_frame, text="Parity:", font=('arial', 10))
        lbl_psw4602_parity.grid(row=4, column=7, padx=5, pady=5)
        psw4602_parity_selection = ttk.Combobox(mte_source_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(psw4602_parity_selection)
        psw4602_parity_selection.configure(state="disabled", width=7)
        psw4602_parity_selection.grid(row=4, column=8, padx=5, pady=5)

        # PSW4602 Stop Bits Drop Down
        lbl_psw4602_sb = ttk.Label(mte_source_frame, text="Stop Bits:", font=('arial', 10))
        lbl_psw4602_sb.grid(row=4, column=9, padx=5, pady=5)
        psw4602_stopbit_selection = ttk.Combobox(mte_source_frame, values=[" ", "1", "2"])
        acc.always_active_style(psw4602_stopbit_selection)
        psw4602_stopbit_selection.configure(state="disabled", width=7)
        psw4602_stopbit_selection.grid(row=4, column=10, padx=5, pady=5)

        # Thermotron Row
        tse_300_checkvar = IntVar()
        tse_300_checkvar.set(0)
        tse_300_chkbtn = ttk.Checkbutton(mte_source_frame, variable=tse_300_checkvar, text="Thermotron", offvalue=0,
                                         onvalue=1)
        tse_300_chkbtn.configure(state="disabled")
        tse_300_chkbtn.grid(row=5, column=0, padx=5, pady=5)

        # Thermotron Com Port Entry
        lbl_tse_300_com = ttk.Label(mte_source_frame, text="Comm. Port:", font=('arial', 10))
        lbl_tse_300_com.grid(row=5, column=1, padx=5, pady=5)
        tse_300_com_value = ttk.Entry(mte_source_frame, state="disabled")
        tse_300_com_value.config(width=3)
        tse_300_com_value.grid(row=5, column=2, padx=5, pady=5)

        # Thermotron Baud Rate Drop Down
        lbl_tse_300_baud = ttk.Label(mte_source_frame, text="Baud Rate:", font=('arial', 10))
        lbl_tse_300_baud.grid(row=5, column=3, padx=5, pady=5)
        tse_300_baud_selection = ttk.Combobox(mte_source_frame, values=[" ", "300", "600", "1200", "1800", "2400",
                                                                        "4800", "7200", "9600", "14400", "19200",
                                                                        "38400", "57600", "115200", "230400",
                                                                        "460800", "921600"])
        acc.always_active_style(tse_300_baud_selection)
        tse_300_baud_selection.configure(state="disabled", width=8)
        tse_300_baud_selection.grid(row=5, column=4, padx=5, pady=5)

        # Thermotron Data Bits Drop Down
        lbl_tse_300_db = ttk.Label(mte_source_frame, text="Data Bits:", font=('arial', 10))
        lbl_tse_300_db.grid(row=5, column=5, padx=5, pady=5)
        tse_300_db_selection = ttk.Combobox(mte_source_frame, values=[" ", "7", "8"])
        acc.always_active_style(tse_300_db_selection)
        tse_300_db_selection.configure(state="disabled", width=3)
        tse_300_db_selection.grid(row=5, column=6, padx=5, pady=5)

        # Thermotron Parity Drop Down
        lbl_tse_300_parity = ttk.Label(mte_source_frame, text="Parity:", font=('arial', 10))
        lbl_tse_300_parity.grid(row=5, column=7, padx=5, pady=5)
        tse_300_parity_selection = ttk.Combobox(mte_source_frame, values=[" ", "Even", "Odd", "None", "Mark", "Space"])
        acc.always_active_style(tse_300_parity_selection)
        tse_300_parity_selection.configure(state="disabled", width=7)
        tse_300_parity_selection.grid(row=5, column=8, padx=5, pady=5)

        # Thermotron Stop Bits Drop Down
        lbl_tse_300_sb = ttk.Label(mte_source_frame, text="Stop Bits:", font=('arial', 10))
        lbl_tse_300_sb.grid(row=5, column=9, padx=5, pady=5)
        tse_300_stopbit_selection = ttk.Combobox(mte_source_frame, values=[" ", "1", "2"])
        acc.always_active_style(tse_300_stopbit_selection)
        tse_300_stopbit_selection.configure(state="disabled", width=7)
        tse_300_stopbit_selection.grid(row=5, column=10, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_da_mte_selection = ttk.Button(da_mte_selection_frame, text="Back", width=20,
                                                  command=lambda: self.data_acq_dut_information(da_mte_selection))
        btn_backout_da_mte_selection.bind("<Return>", lambda event: self.data_acq_dut_information(da_mte_selection))
        btn_backout_da_mte_selection.grid(row=5, column=0, columnspan=2, padx=5, pady=5)

        btn_config_channels = ttk.Button(da_mte_selection_frame, text="Configure Channels", width=20,
                                         command=lambda: self.__init__())
        btn_config_channels.bind("<Return>", lambda event: self.__init__())
        btn_config_channels.config(state="disabled")
        btn_config_channels.grid(row=5, column=2, columnspan=2, padx=5, pady=5)

        btn_setup_profile = ttk.Button(da_mte_selection_frame, text="Next", width=20,
                                       command=lambda: self.update_data_acq_serial_comm())
        btn_setup_profile.bind("<Return>", lambda event: self.update_data_acq_serial_comm())
        btn_setup_profile.grid(row=5, column=4, columnspan=2, padx=5, pady=5)

    # -----------------------------------------------------------------------------#

    # Function created to enable/disable entry fields based on previous user input
    def process_data_acq_chkbtn(self):
        self.__init__()

        if mc6_checkvar.get() == 1:
            mc6_com_value.configure(state="active")
            mc6_baud_selection.configure(state="active")
            mc6_db_selection.configure(state="active")
            mc6_parity_selection.configure(state="active")
            mc6_stopbit_selection.configure(state="active")
        else:
            mc6_com_value.configure(state="disabled")
            mc6_baud_selection.configure(state="disabled")
            mc6_db_selection.configure(state="disabled")
            mc6_parity_selection.configure(state="disabled")
            mc6_stopbit_selection.configure(state="disabled")

    # -----------------------------------------------------------------------------#

    # Function to update serial communications for reference standards selected
    def update_data_acq_serial_comm(self):
        self.__init__()

        if mc6_checkvar.get() == 1:
            LIMSVarConfig.mc6_com_port = mc6_com_value.get()
            LIMSVarConfig.mc6_baud = mc6_baud_selection.get()
            LIMSVarConfig.mc6_data_bits = mc6_db_selection.get()
            LIMSVarConfig.mc6_parity = mc6_parity_selection.get()
            LIMSVarConfig.mc6_stop_bit = mc6_stopbit_selection.get()
        else:
            self.__init__()

        self.data_acq_test_profile(da_mte_selection)

    # -----------------------------------------------------------------------------#

    # Function to allow user to create test profile/pattern for data acquisition
    def data_acq_test_profile(self, window):
        self.__init__()
        window.withdraw()

        LIMSVarConfig.data_acquisition_test_step = []
        LIMSVarConfig.data_acquisition_set_point = []
        LIMSVarConfig.data_acquisition_soak_time = []
        LIMSVarConfig.data_acquisition_sample_rate = []
        LIMSVarConfig.data_acquisition_temperature_set_point = []

        global da_profile_creation

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        da_profile_creation = Toplevel()
        da_profile_creation.title("Data Acquisition - Test Profile")
        da_profile_creation.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 620
        height = 360
        screen_width = da_profile_creation.winfo_screenwidth()
        screen_height = da_profile_creation.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        da_profile_creation.geometry("%dx%d+%d+%d" % (width, height, x, y))
        da_profile_creation.focus_force()
        da_profile_creation.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(da_profile_creation))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(da_profile_creation, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(da_profile_creation)),
                                           ("D.A. - M&TE Selection",
                                            lambda: self.data_acq_mte_selection(da_profile_creation)),
                                           ("Logout", lambda: acc.software_signout(da_profile_creation)),
                                           ("Quit", lambda: acc.software_close(da_profile_creation))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.data_acq_profile_creation_help())])

        # ........................Canvas & Frame Creation........................#

        da_profile_creation_frame = LabelFrame(da_profile_creation, text="Data Acquisition - Test Profile Creation",
                                               relief=SOLID, bd=1, labelanchor="n")
        da_profile_creation_frame.grid(row=2, column=0, rowspan=4, columnspan=6, padx=5, pady=5)

        da_profile_canvas = Canvas(da_profile_creation_frame)
        da_profile_canvas.grid(row=0, column=0, padx=5, pady=5)

        da_profile_full_table = LabelFrame(da_profile_canvas, relief=SOLID, bd=1)

        # ...............................Scrollbar...............................#

        da_profile_scrollbar = ttk.Scrollbar(da_profile_creation_frame, orient=VERTICAL,
                                             command=da_profile_canvas.yview)
        da_profile_scrollbar.grid(row=0, column=1, sticky=NS)
        da_profile_canvas.configure(yscrollcommand=da_profile_scrollbar.set)

        # .......................Labels, Dropdowns & Entries.....................#

        # Load Existing Profile
        btn_load_da_test_profile = ttk.Button(da_profile_creation, text="Open Existing Profile", width=20,
                                              command=lambda: self.__init__(), state="disabled")
        btn_load_da_test_profile.bind("<Return>", lambda event: self.__init__())
        btn_load_da_test_profile.grid(row=0, column=0, padx=5, pady=5)

        lbl_dummy = ttk.Label(da_profile_creation, text="")
        lbl_dummy.config(width=4)
        lbl_dummy.grid(row=0, column=1, padx=5, pady=5)

        # Profile Name
        lbl_profile_name = ttk.Label(da_profile_creation, text="Profile Name:")
        lbl_profile_name.grid(row=0, column=2, padx=5, pady=5)
        profile_name_value = ttk.Entry(da_profile_creation, state="disabled")
        profile_name_value.grid(row=0, column=3, padx=5, pady=5)

        # Number of Loops Selection
        lbl_number_of_loops = ttk.Label(da_profile_creation, text="# of Loops:")
        lbl_number_of_loops.grid(row=0, column=4, padx=5, pady=5)
        number_of_loops_selection = ttk.Combobox(da_profile_creation, values=["0", "1", "2", "3", "4", "5", "6", "7",
                                                                              "8", "9", "10", "15", "20"])
        acc.always_active_style(number_of_loops_selection)
        number_of_loops_selection.configure(state="active", width=18)
        number_of_loops_selection.grid(row=0, column=5, padx=5, pady=5)

        test_profile_rows = 21
        test_profile_columns = 5
        test_profile_rows_display = 10
        test_profile_columns_display = 5

        # Table Creation
        for i in range(0, test_profile_rows + 1):
            for j in range(test_profile_columns):
                da_profile_table = ttk.Label(da_profile_full_table, text=" ", anchor="n")
                da_profile_table.grid(row=i, column=j)

        for i in range(0, test_profile_rows + 1):
            da_profile_table = ttk.Label(da_profile_full_table, text=i, anchor="n")
            da_profile_table.config(relief=SOLID, width=10)
            da_profile_table.grid(row=i, column=0, sticky="n"+"e"+"s"+"w")
            LIMSVarConfig.data_acquisition_test_step.append(i+1)

        # Table Headers
        lbl_test_profile_test_step = ttk.Label(da_profile_full_table, text="Test Step", anchor="n")
        lbl_test_profile_test_step.config(relief=SOLID)
        lbl_test_profile_test_step.grid(row=0, column=0, sticky="news")

        lbl_test_profile_set_point = ttk.Label(da_profile_full_table, text="Set Point", anchor="n")
        lbl_test_profile_set_point.config(relief=SOLID)
        lbl_test_profile_set_point.grid(row=0, column=1, sticky="news")

        lbl_test_profile_temperature = ttk.Label(da_profile_full_table, text="Temperature Set Point", anchor="n")
        lbl_test_profile_temperature.config(relief=SOLID)
        lbl_test_profile_temperature.grid(row=0, column=2, sticky="news")

        lbl_test_profile_soak_time = ttk.Label(da_profile_full_table, text="Soak Time (Min.)", anchor="n")
        lbl_test_profile_soak_time.config(relief=SOLID)
        lbl_test_profile_soak_time.grid(row=0, column=3, sticky="news")

        lbl_test_profile_sample = ttk.Label(da_profile_full_table, text="Sample Rate (Sec.)", anchor="n")
        lbl_test_profile_sample.config(relief=SOLID)
        lbl_test_profile_sample.grid(row=0, column=4, sticky="news")

        for i in range(1, test_profile_rows + 1):
            set_point_values = ttk.Entry(da_profile_full_table)
            set_point_values.grid(row=i, column=1)
            LIMSVarConfig.data_acquisition_set_point.append(set_point_values)
            temperature_set_point_values = ttk.Entry(da_profile_full_table, state="disabled")
            temperature_set_point_values.grid(row=i, column=2)
            LIMSVarConfig.data_acquisition_temperature_set_point.append(temperature_set_point_values)
            soak_time_values = ttk.Entry(da_profile_full_table)
            soak_time_values.grid(row=i, column=3)
            LIMSVarConfig.data_acquisition_soak_time.append(soak_time_values)
            sampling_rate_values = ttk.Entry(da_profile_full_table)
            sampling_rate_values.grid(row=i, column=4)
            LIMSVarConfig.data_acquisition_sample_rate.append(sampling_rate_values)

        # Creation of window for labels to exist in
        da_profile_canvas.create_window((0, 0), window=da_profile_full_table, anchor=NW)

        # Required to make bbox information available
        da_profile_full_table.update_idletasks()
        # Get bound box of canvas with labels
        bbox = da_profile_canvas.bbox(ALL)

        # Define scrollable region as entire canvas with only the desired number of rows and columns displayed
        w, h = bbox[2] - bbox[1], bbox[3] - bbox[1]
        dw, dh = int((w/test_profile_columns) * test_profile_columns_display), int((h/test_profile_rows) *
                                                                                   test_profile_rows_display)
        da_profile_canvas.configure(scrollregion=bbox, width=dw, height=dh)

        # ...............................Buttons..................................#

        btn_backout_da_test_profile = ttk.Button(da_profile_creation, text="Back", width=20,
                                                 command=lambda: self.data_acq_mte_selection(da_profile_creation))
        btn_backout_da_test_profile.config(width=20)
        btn_backout_da_test_profile.bind("<Return>", lambda event: self.data_acq_mte_selection(da_profile_creation))
        btn_backout_da_test_profile.grid(row=6, column=0, columnspan=2, padx=5, pady=5)

        btn_save_da_test_profile = ttk.Button(da_profile_creation, text="Save Profile", width=20,
                                              command=lambda: self.__init__(), state="disabled")
        btn_save_da_test_profile.config(width=20)
        btn_save_da_test_profile.bind("<Return>", lambda event: self.__init__())
        btn_save_da_test_profile.grid(row=6, column=2, columnspan=2, padx=5, pady=5)

        btn_perform_data_acq = ttk.Button(da_profile_creation, text="Begin Test", width=20,
                                          command=lambda: self.perform_data_acquisition(da_profile_creation))
        btn_perform_data_acq.config(width=20)
        btn_perform_data_acq.bind("<Return>", lambda event: self.perform_data_acquisition(da_profile_creation))
        btn_perform_data_acq.grid(row=6, column=4, columnspan=2, padx=5, pady=5)

# -----------------------------------------------------------------------------#

    # Function to allow user to create test profile/pattern for data acquisition
    def perform_data_acquisition(self, window):
        self.__init__()
        window.withdraw()

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        automated_data_acquisition_window = Toplevel()
        automated_data_acquisition_window.title("Data Acquisition - Automated")
        automated_data_acquisition_window.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 720
        height = 360
        screen_width = automated_data_acquisition_window.winfo_screenwidth()
        screen_height = automated_data_acquisition_window.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        automated_data_acquisition_window.geometry("%dx%d+%d+%d" % (width, height, x, y))
        automated_data_acquisition_window.focus_force()
        automated_data_acquisition_window.protocol("WM_DELETE_WINDOW",
                                                   lambda: acc.on_exit(automated_data_acquisition_window))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(automated_data_acquisition_window, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(automated_data_acquisition_window)),
                                           ("D.A. - Test Profile",
                                            lambda: self.data_acq_test_profile(automated_data_acquisition_window)),
                                           ("Logout", lambda: acc.software_signout(automated_data_acquisition_window)),
                                           ("Quit", lambda: acc.software_close(automated_data_acquisition_window))])
        menubar.add_menu("Help", commands=[("Help", lambda: self.__init__())])
        menubar.add_menu("Abort", commands=[("Abort", lambda: self.__init__())])

        # ........................Canvas & Frame Creation........................#

        automated_da_frame = LabelFrame(automated_data_acquisition_window, text="Data Acquisition",
                                        relief=SOLID, bd=1, labelanchor="n")
        automated_da_frame.grid(row=0, column=0, rowspan=4, columnspan=6, padx=5, pady=5)

        # .........................Labels and Entries............................#

        automated_da_textbox = Text(automated_da_frame, wrap="none", width=80, height=12)
        automated_da_vscroll = ttk.Scrollbar(automated_da_frame, orient='vertical', command=automated_da_textbox.yview)
        automated_da_hscroll = ttk.Scrollbar(automated_da_frame, orient='horizontal',
                                             command=automated_da_textbox.xview)
        automated_da_textbox.configure(font=('arial', 11))
        automated_da_textbox['yscroll'] = automated_da_vscroll.set
        automated_da_textbox['xscroll'] = automated_da_hscroll.set
        automated_da_vscroll.pack(side=RIGHT, fill=Y, padx=5)
        automated_da_hscroll.pack(side=BOTTOM, fill=X, pady=5)
        automated_da_textbox.pack(fill=BOTH, expand=Y, padx=5, pady=5)

        # .............................Progress Bar..............................#

        prgrbar_automated_da = ttk.Progressbar(automated_data_acquisition_window, orient='horizontal',
                                               length=100, mode='determinate')
        prgrbar_automated_da.grid(row=5, column=0, columnspan=6, padx=5, pady=8)
        prgrbar_automated_da['value'] = 0
