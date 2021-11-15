"""
LIMSDataImport is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of
this module is to import data obtained from calibration systems that create .dat files or excel files without
having to rewrite data obtained.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import csv
import os
import os.path

import LIMSVarConfig
import win32com.client
from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk
from tkinter import filedialog


class AppDataImportModule:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # ==========================DATA IMPORT FUNCTIONALITY==========================#

    # Data Import System Selection Window
    def data_import_selection(self, window):
        self.window = window
        window.withdraw()

        global DataImportSelectionWindow

        from LIMSHomeWindow import AppCommonCommands
        from LIMSCertCreation import AppCalibrationModule
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        acm = AppCalibrationModule()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        DataImportSelectionWindow = Toplevel()
        DataImportSelectionWindow.title("Data Import System Selection")
        DataImportSelectionWindow.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 310
        height = 200
        screen_width = DataImportSelectionWindow.winfo_screenwidth()
        screen_height = DataImportSelectionWindow.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        DataImportSelectionWindow.geometry("%dx%d+%d+%d" % (width, height, x, y))
        DataImportSelectionWindow.focus_force()
        DataImportSelectionWindow.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(DataImportSelectionWindow))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(DataImportSelectionWindow, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(DataImportSelectionWindow)),
                                           ("Calibration Reference Standards",
                                            lambda: acm.calibration_standards_selection(DataImportSelectionWindow)),
                                           ("Logout", lambda: acc.software_signout(DataImportSelectionWindow)),
                                           ("Quit", lambda: acc.software_close(DataImportSelectionWindow))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.data_import_selection_help())])

        # ..........................Frame Creation...............................#

        data_import_system_frame = LabelFrame(DataImportSelectionWindow,
                                              text="Select System Used to Collect Calibration Data", relief=SOLID,
                                              bd=1, labelanchor="n")
        data_import_system_frame.grid(row=0, column=0, rowspan=4, columnspan=2, padx=8, pady=5)

        # ...............................Buttons..................................#

        btn_fluke2465_a_system = ttk.Button(data_import_system_frame, text="Fluke 2465/8A-754 Piston Gauge System",
                                            command=lambda: self.fluke_24658a_afal_as_and_al_query(DataImportSelectionWindow))
        btn_fluke2465_a_system.bind("<Return>",
                                    lambda event: self.fluke_24658a_afal_as_and_al_query(DataImportSelectionWindow))
        btn_fluke2465_a_system.grid(row=1, column=0, columnspan=2, pady=5, padx=5)
        btn_fluke2465_a_system.config(width=45)

        btn_fluke_molbox_system = ttk.Button(data_import_system_frame, text="Fluke Molbox Gas Flow System",
                                             command=lambda: self.fluke_molbox_afal_as_and_al_query(DataImportSelectionWindow))
        btn_fluke_molbox_system.bind("<Return>",
                                     lambda event: self.fluke_molbox_afal_as_and_al_query(DataImportSelectionWindow))
        btn_fluke_molbox_system.grid(row=2, column=0, columnspan=2, pady=5, padx=5)
        btn_fluke_molbox_system.config(width=45)

        btn_sonic_nozzle_system = ttk.Button(data_import_system_frame, text="Sonic Nozzle System",
                                             command=lambda: self.sonic_nozzle_data_import())
        btn_sonic_nozzle_system.bind("<Return>", lambda event: self.sonic_nozzle_data_import())
        btn_sonic_nozzle_system.grid(row=3, column=0, columnspan=2, pady=5, padx=5)
        btn_sonic_nozzle_system.config(width=45)

        btn_data_import_back_out = ttk.Button(DataImportSelectionWindow, text="Back",
                                              command=lambda: acm.calibration_standards_selection(DataImportSelectionWindow))
        btn_data_import_back_out.bind("<Return>",
                                      lambda event: acm.calibration_standards_selection(DataImportSelectionWindow))
        btn_data_import_back_out.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        btn_data_import_back_out.config(width=20)

    # -----------------------------------------------------------------------#

    # Command to extract calibration data from Fluke 2465/2468A .dat File and store to Arrays for certificate
    def fluke_24658a_afal_as_and_al_query(self, window):
        self.window = window
        window.withdraw()

        global dwtester_data_set_query_window, dwtester_data_set_selection, dwtester_mode_selection

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        dwtester_data_set_query_window = Toplevel()
        dwtester_data_set_query_window.title("Fluke 2465/8A - Data Set Selection")
        dwtester_data_set_query_window.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 340
        height = 170
        screen_width = dwtester_data_set_query_window.winfo_screenwidth()
        screen_height = dwtester_data_set_query_window.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        dwtester_data_set_query_window.geometry("%dx%d+%d+%d" % (width, height, x, y))
        dwtester_data_set_query_window.focus_force()
        dwtester_data_set_query_window.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(dwtester_data_set_query_window))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(dwtester_data_set_query_window, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(dwtester_data_set_query_window)),
                                           ("Data Import System Selection",
                                            lambda: self.data_import_selection(dwtester_data_set_query_window)),
                                           ("Logout", lambda: acc.software_signout(dwtester_data_set_query_window)),
                                           ("Quit", lambda: acc.software_close(dwtester_data_set_query_window))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.dwtester_data_set_query_help())])

        # ..........................Frame Creation...............................#

        dwtester_data_set_query_frame = LabelFrame(dwtester_data_set_query_window,
                                                   text="Select Data Set Type \n Obtained from Calibration",
                                                   relief=SOLID, bd=1, labelanchor="n")
        dwtester_data_set_query_frame.grid(row=0, column=0, rowspan=2, columnspan=1, padx=8, pady=5)

        dwtester_mode_query_frame = LabelFrame(dwtester_data_set_query_window,
                                               text="Select Mode of Operation \n Used for Calibration",
                                               relief=SOLID, bd=1, labelanchor="n")
        dwtester_mode_query_frame.grid(row=0, column=1, rowspan=2, columnspan=1, padx=8, pady=5)

        # .......................Drop Down Lists................................. #

        dwtester_data_set_selection = ttk.Combobox(dwtester_data_set_query_frame,
                                                   values=[" ", "As Found / As Left", "As Found & As Left",
                                                           "As Found Only", "As Left Only"])
        acc.always_active_style(dwtester_data_set_selection)
        dwtester_data_set_selection.configure(state="active", width=18)
        dwtester_data_set_selection.grid(padx=10, pady=5, row=0, column=1)

        dwtester_mode_selection = ttk.Combobox(dwtester_mode_query_frame,
                                               values=[" ", "Bi-Directional", "Pos., Neg. or Abs."])
        acc.always_active_style(dwtester_mode_selection)
        dwtester_mode_selection.configure(state="active", width=18)
        dwtester_mode_selection.grid(padx=10, pady=5, row=0, column=1, columnspan=2)

        # ...............................Buttons..................................#

        btn_dwtester_data_query_accept = ttk.Button(dwtester_data_set_query_window, text="Open",
                                                    command=lambda: self.fluke_24658a_data_import())
        btn_dwtester_data_query_accept.bind("<Return>", lambda event: self.fluke_24658a_data_import())
        btn_dwtester_data_query_accept.grid(row=3, column=0, columnspan=2, padx=10, pady=5)
        btn_dwtester_data_query_accept.config(width=20)

        btn_dwtester_data_query_back_out = ttk.Button(dwtester_data_set_query_window, text="Back",
                                                      command=lambda: self.data_import_selection(dwtester_data_set_query_window))
        btn_dwtester_data_query_back_out.bind("<Return>",
                                              lambda event: self.data_import_selection(dwtester_data_set_query_window))
        btn_dwtester_data_query_back_out.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        btn_dwtester_data_query_back_out.config(width=20)

    # -----------------------------------------------------------------------#

    # Command to extract calibration data from Fluke 2465/2468A .dat File and store to Arrays for certificate
    def fluke_24658a_data_import(self):

        LIMSVarConfig.imported_data_checker = int(0)
        LIMSVarConfig.imported_dut_reading = float()
        LIMSVarConfig.imported_ref_reading = ""
        LIMSVarConfig.imported_tolerance_reading = ""
        LIMSVarConfig.imported_device_reading_list = []
        LIMSVarConfig.imported_reference_reading_list = []
        LIMSVarConfig.imported_total_error_band_list = []
        LIMSVarConfig.imported_measured_difference_list = []
        LIMSVarConfig.imported_pass_fail_list = []
        LIMSVarConfig.second_imported_data_checker = int(0)
        LIMSVarConfig.second_imported_dut_reading = ""
        LIMSVarConfig.second_imported_ref_reading = ""
        LIMSVarConfig.second_imported_tolerance_reading = ""
        LIMSVarConfig.second_imported_device_reading_list = []
        LIMSVarConfig.second_imported_reference_reading_list = []
        LIMSVarConfig.second_imported_total_error_band_list = []
        LIMSVarConfig.second_imported_measured_difference_list = []
        LIMSVarConfig.second_imported_pass_fail_list = []

        if dwtester_data_set_selection.get() == "As Found / As Left" and \
                dwtester_mode_selection.get() == "Pos., Neg. or Abs.":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your 'As Found'/'As Left' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Found / As Left" and \
                dwtester_mode_selection.get() == "Bi-Directional":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select the Negative 'As Found'/'As Left' \
Fluke 2465/8A Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_2 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select the Positive 'As Found'/'As Left' \
Fluke 2465/8A Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Found & As Left" and \
                dwtester_mode_selection.get() == "Pos., Neg. or Abs.":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your 'As Found' Fluke 2465/8A Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_2 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select Your 'As Left' Fluke 2465/8A Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Found & As Left" and \
                dwtester_mode_selection.get() == "Bi-Directional":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your Negative 'As Found' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_2 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select Your Positive 'As Found' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_3 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select Your Negative 'As Left' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_4 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select Your Positive 'As Left' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Found Only" and \
                dwtester_mode_selection.get() == "Pos., Neg. or Abs.":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your 'As Found' Fluke 2465/8A Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Found Only" and \
                dwtester_mode_selection.get() == "Bi-Directional":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your Negative 'As Found' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_2 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select Your Positive 'As Found' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Left Only" and \
                dwtester_mode_selection.get() == "Pos., Neg. or Abs.":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your 'As Left' Fluke 2465/8A Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif dwtester_data_set_selection.get() == "As Left Only" and \
                dwtester_mode_selection.get() == "Bi-Directional":

            fluke_2465_8a_search = filedialog.askopenfilename(initialdir="/",
                                                                title="Select Your Negative 'As Left' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_2465_8a_search_2 = filedialog.askopenfilename(initialdir="/",
                                                                  title="Select Your Positive 'As Left' Fluke 2465/8A \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        else:

            tm.showerror("Invalid Selection", "Please select a data set type from the provided options!")

        if (dwtester_data_set_selection.get() != "" or dwtester_data_set_selection.get() != " ") and \
                (dwtester_mode_selection.get() != "" or dwtester_mode_selection.get() != " "):

            with open(fluke_2465_8a_search.strip('.dat') + '.csv', 'wb') as f_2465_output_file:
                with open(fluke_2465_8a_search) as f_2465_file:
                    lines = f_2465_file.readlines()
                    newlines = []
                    for line in lines:
                        newline = line.strip().split(';')
                        newlines.append(newline)
                f_2465_file_writer = csv.writer(f_2465_output_file)
                f_2465_file_writer.writerows(newlines)

                f_2465_saved_file = os.path.split(fluke_2465_8a_search.strip('.dat'))[-1]

            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            work_book = excel.Workbooks.Open(fluke_2465_8a_search.strip('.dat') + '.csv')
            sheet = work_book.Sheets(f_2465_saved_file)

            # Extract Data Collected from Fluke 2465/8A System
            for i in range(79, 100):
                if sheet.Cells(i, 1).Value is not None:
                    LIMSVarConfig.imported_ref_reading = sheet.Cells(i, 11).Value  # Reference Readings
                    LIMSVarConfig.imported_dut_reading = sheet.Cells(i, 12).Value  # Device Under Test Readings
                    LIMSVarConfig.imported_tolerance_reading = sheet.Cells(i, 14).Value  # DUT Tolerance
                    LIMSVarConfig.imported_device_reading_list.append(LIMSVarConfig.imported_dut_reading)
                    LIMSVarConfig.imported_reference_reading_list.append(LIMSVarConfig.imported_ref_reading)
                    LIMSVarConfig.imported_total_error_band_list.append(LIMSVarConfig.imported_tolerance_reading)
                    LIMSVarConfig.imported_measured_difference_list.append(LIMSVarConfig.imported_dut_reading -
                                                                           LIMSVarConfig.imported_ref_reading)
                    i += 1
                else:
                    break

            for i in range(0, len(LIMSVarConfig.imported_total_error_band_list)):
                if -float(LIMSVarConfig.imported_total_error_band_list[i]) <= \
                        float(LIMSVarConfig.imported_measured_difference_list[i]) <= \
                        float(LIMSVarConfig.imported_total_error_band_list[i]):
                    LIMSVarConfig.imported_pass_fail_list.append("Pass")
                else:
                    LIMSVarConfig.imported_pass_fail_list.append("Fail")

            # Create Measured Difference List and Pass Fail Criteria

            work_book.Close(True)

            if (dwtester_data_set_selection.get() == "As Found / As Left" and
                    dwtester_mode_selection.get() == "Bi-Directional") or \
                    (dwtester_data_set_selection.get() == "As Found & As Left" and
                     dwtester_mode_selection.get() == "Pos., Neg. or Abs.") or \
                    (dwtester_data_set_selection.get() == "As Found" and
                     dwtester_mode_selection.get() == "Bi-Directional") or \
                    (dwtester_data_set_selection.get() == "As Left" and
                     dwtester_mode_selection.get() == "Bi-Directional"):

                with open(fluke_2465_8a_search_2.strip('.dat') + '.csv', 'wb') as f_2465_output_file_2:
                    with open(fluke_2465_8a_search_2) as f_2465_file_2:
                        lines_2 = f_2465_file_2.readlines()
                        newlines_2 = []
                        for line_2 in lines_2:
                            newline_2 = line_2.strip().split(';')
                            newlines_2.append(newline_2)
                    f_2465_file_writer_2 = csv.writer(f_2465_output_file_2)
                    f_2465_file_writer_2.writerows(newlines_2)

                    f_2465_saved_file_2 = os.path.split(fluke_2465_8a_search_2.strip('.dat'))[-1]

                excel = win32com.client.dynamic.Dispatch("Excel.Application")
                work_book = excel.Workbooks.Open(fluke_2465_8a_search_2.strip('.dat') + '.csv')
                sheet = work_book.Sheets(f_2465_saved_file_2)

                # Extract Data Collected from Fluke 2468A System
                for i in range(79, 100):
                    if sheet.Cells(i, 1).Value is not None:
                        LIMSVarConfig.second_imported_ref_reading = sheet.Cells(i, 11).Value  # Reference Readings
                        LIMSVarConfig.second_imported_dut_reading = sheet.Cells(i,
                                                                                12).Value  # Device Under Test Readings
                        LIMSVarConfig.second_imported_tolerance_reading = sheet.Cells(i, 14).Value  # DUT Tolerance
                        LIMSVarConfig.second_imported_device_reading_list.append(
                            LIMSVarConfig.second_imported_dut_reading)
                        LIMSVarConfig.second_imported_reference_reading_list.append(
                            LIMSVarConfig.second_imported_ref_reading)
                        LIMSVarConfig.second_imported_total_error_band_list.append(
                            LIMSVarConfig.second_imported_tolerance_reading)
                        LIMSVarConfig.second_imported_measured_difference_list.append(
                            LIMSVarConfig.second_imported_dut_reading -
                            LIMSVarConfig.second_imported_ref_reading)
                        i += 1
                    else:
                        break

                for i in range(0, len(LIMSVarConfig.second_imported_total_error_band_list)):
                    if -float(LIMSVarConfig.second_imported_total_error_band_list[i]) <= \
                            float(LIMSVarConfig.second_imported_measured_difference_list[i]) <= \
                            float(LIMSVarConfig.second_imported_total_error_band_list[i]):
                        LIMSVarConfig.second_imported_pass_fail_list.append("Pass")
                    else:
                        LIMSVarConfig.second_imported_pass_fail_list.append("Fail")

                # Create Measured Difference List and Pass Fail Criteria

                work_book.Close(True)

                LIMSVarConfig.second_imported_data_checker = int(1)

            else:

                self.__init__()

            certificate_query = tm.askyesno("Complete Calibration?", "Would you like to generate a Certificate of \
Calibration for the device under test? If you would like to select a different file or cancel out of the \
process, click the 'No' button.")

            LIMSVarConfig.imported_data_checker = int(1)

            if certificate_query is True:
                from LIMSCertCreation import AppCalibrationModule
                acm = AppCalibrationModule()
                acm.complete_calibration_process_step_two(dwtester_data_set_query_window)

            elif certificate_query is False:
                LIMSVarConfig.clear_imported_data_variables()

        else:
            self.__init__()

    # -----------------------------------------------------------------------#

    # Command to request type of data set user is importing

    def fluke_molbox_afal_as_and_al_query(self, window):
        self.window = window
        window.withdraw()

        global data_set_query_window, data_set_selection

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        data_set_query_window = Toplevel()
        data_set_query_window.title("Fluke Molbox - Data Set Selection")
        data_set_query_window.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 325
        height = 120
        screen_width = data_set_query_window.winfo_screenwidth()
        screen_height = data_set_query_window.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        data_set_query_window.geometry("%dx%d+%d+%d" % (width, height, x, y))
        data_set_query_window.focus_force()
        data_set_query_window.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(data_set_query_window))

        # .........................Menu Bar Creation..................................#

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(data_set_query_window, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(data_set_query_window)),
                                           ("Data Import System Selection",
                                            lambda: self.data_import_selection(data_set_query_window)),
                                           ("Logout", lambda: acc.software_signout(data_set_query_window)),
                                           ("Quit", lambda: acc.software_close(data_set_query_window))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.data_set_query_help())])

        # ..........................Frame Creation...............................#

        data_set_query_frame = LabelFrame(data_set_query_window,
                                          text="Select Data Set Type Obtained from Calibration",
                                          relief=SOLID, bd=1, labelanchor="n")
        data_set_query_frame.grid(row=0, column=0, rowspan=4, columnspan=2, padx=8, pady=5)

        # .......................Drop Down Lists................................. #

        data_set_selection = ttk.Combobox(data_set_query_frame, values=[" ", "As Found / As Left", "As Found & As Left",
                                                                        "As Found Only", "As Left Only"])
        acc.always_active_style(data_set_selection)
        data_set_selection.configure(state="active", width=18)
        data_set_selection.grid(padx=10, pady=5, row=0, column=0, columnspan=2)

        # ...............................Buttons..................................#

        btn_data_query_accept = ttk.Button(data_set_query_frame, text="Open",
                                           command=lambda: self.fluke_molbox_data_import())
        btn_data_query_accept.bind("<Return>", lambda event: self.fluke_molbox_data_import())
        btn_data_query_accept.grid(row=0, column=2, columnspan=2, padx=10, pady=5)
        btn_data_query_accept.config(width=20)

        btn_data_query_back_out = ttk.Button(data_set_query_window, text="Back",
                                             command=lambda: self.data_import_selection(data_set_query_window))
        btn_data_query_back_out.bind("<Return>", lambda event: self.data_import_selection(data_set_query_window))
        btn_data_query_back_out.grid(row=4, column=0, columnspan=2, padx=5, pady=5)
        btn_data_query_back_out.config(width=20)

    # -----------------------------------------------------------------------#

    # Command to extract calibration data from Fluke Molbox .dat File and store to Arrays for certificate
    def fluke_molbox_data_import(self):

        LIMSVarConfig.imported_data_checker = int(0)
        LIMSVarConfig.imported_dut_reading = ""
        LIMSVarConfig.imported_ref_reading = ""
        LIMSVarConfig.imported_tolerance_reading = ""
        LIMSVarConfig.imported_device_reading_list = []
        LIMSVarConfig.imported_reference_reading_list = []
        LIMSVarConfig.imported_total_error_band_list = []
        LIMSVarConfig.imported_measured_difference_list = []
        LIMSVarConfig.imported_pass_fail_list = []
        LIMSVarConfig.second_imported_data_checker = int(0)
        LIMSVarConfig.second_imported_dut_reading = ""
        LIMSVarConfig.second_imported_ref_reading = ""
        LIMSVarConfig.second_imported_tolerance_reading = ""
        LIMSVarConfig.second_imported_device_reading_list = []
        LIMSVarConfig.second_imported_reference_reading_list = []
        LIMSVarConfig.second_imported_total_error_band_list = []
        LIMSVarConfig.second_imported_measured_difference_list = []
        LIMSVarConfig.second_imported_pass_fail_list = []

        if data_set_selection.get() == "As Found / As Left":

            fluke_molbox_search = filedialog.askopenfilename(initialdir="/",
                                                               title="Select Your 'As Found'/'As Left' Fluke Molbox \
Data File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif data_set_selection.get() == "As Found & As Left":

            fluke_molbox_search = filedialog.askopenfilename(initialdir="/",
                                                               title="Select Your 'As Found' Fluke Molbox Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

            fluke_molbox_search_2 = filedialog.askopenfilename(initialdir="/",
                                                               title="Select Your 'As Left' Fluke Molbox Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif data_set_selection.get() == "As Found Only":

            fluke_molbox_search = filedialog.askopenfilename(initialdir="/",
                                                               title="Select Your 'As Found' Fluke Molbox Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        elif data_set_selection.get() == "As Left Only":

            fluke_molbox_search = filedialog.askopenfilename(initialdir="/",
                                                               title="Select Your 'As Left' Fluke Molbox Data \
File", filetypes=(("data files", "*.dat"), ("all files", "*.*")))

        else:

            tm.showerror("Invalid Selection", "Please select a data set type from the provided options!")

        if data_set_selection.get() == "As Found / As Left" or data_set_selection.get() == "As Found Only" or \
                data_set_selection.get() == "As Found & As Left" or data_set_selection.get() == "As Left Only":

            with open(fluke_molbox_search.strip('.dat') + '.csv', 'wb') as f_m_output_file:
                with open(fluke_molbox_search) as f_m_file:
                    lines = f_m_file.readlines()
                    newlines = []
                    for line in lines:
                        newline = line.strip().split(';')
                        newlines.append(newline)
                f_m_file_writer = csv.writer(f_m_output_file)
                f_m_file_writer.writerows(newlines)

            asfound_asleft_file = os.path.split(fluke_molbox_search.strip('.dat'))[-1]

            excel = win32com.client.dynamic.Dispatch("Excel.Application")
            work_book = excel.Workbooks.Open(fluke_molbox_search.strip('.dat') + '.csv')
            sheet = work_book.Sheets(asfound_asleft_file)

            # Extract Data Collected from Fluke Molbox System
            for i in range(43, 63):
                if sheet.Cells(i, 11).Value is not None:
                    LIMSVarConfig.imported_ref_reading = sheet.Cells(i, 11).Value  # Reference Readings
                    LIMSVarConfig.imported_dut_reading = sheet.Cells(i, 12).Value  # Device Under Test Readings
                    LIMSVarConfig.imported_tolerance_reading = sheet.Cells(i, 14).Value  # DUT Tolerance
                    LIMSVarConfig.imported_device_reading_list.append(LIMSVarConfig.imported_dut_reading)
                    LIMSVarConfig.imported_reference_reading_list.append(LIMSVarConfig.imported_ref_reading)
                    LIMSVarConfig.imported_total_error_band_list.append(LIMSVarConfig.imported_tolerance_reading)
                    LIMSVarConfig.imported_measured_difference_list.append(LIMSVarConfig.imported_dut_reading -
                                                                           LIMSVarConfig.imported_ref_reading)
                    i += 1
                else:
                    break

            for i in range(0, len(LIMSVarConfig.imported_total_error_band_list)):
                if -float(LIMSVarConfig.imported_total_error_band_list[i]) <= \
                        float(LIMSVarConfig.imported_measured_difference_list[i]) <= \
                        float(LIMSVarConfig.imported_total_error_band_list[i]):
                    LIMSVarConfig.imported_pass_fail_list.append("Pass")
                else:
                    LIMSVarConfig.imported_pass_fail_list.append("Fail")

            # Create Measured Difference List and Pass Fail Criteria

            work_book.Close(True)

            if data_set_selection.get() == "As Found & As Left":
                with open(fluke_molbox_search_2.strip('.dat') + '.csv', 'wb') as f_m_output_file_2:
                    with open(fluke_molbox_search_2) as f_m_file_2:
                        lines_2 = f_m_file_2.readlines()
                        newlines_2 = []
                        for line in lines_2:
                            newline_2 = line.strip().split(';')
                            newlines_2.append(newline_2)
                    f_m_file_2_writer = csv.writer(f_m_output_file_2)
                    f_m_file_2_writer.writerows(newlines_2)

                asfound_asleft_file_2 = os.path.split(fluke_molbox_search_2.strip('.dat'))[-1]

                excel = win32com.client.dynamic.Dispatch("Excel.Application")
                work_book = excel.Workbooks.Open(fluke_molbox_search_2.strip('.dat') + '.csv')
                sheet = work_book.Sheets(asfound_asleft_file_2)

                # Extract Data Collected from Fluke Molbox System
                for i in range(43, 63):
                    if sheet.Cells(i, 11).Value is not None:
                        LIMSVarConfig.second_imported_ref_reading = sheet.Cells(i, 11).Value  # Reference Readings
                        LIMSVarConfig.second_imported_dut_reading = sheet.Cells(i, 12).Value  # Device Under Test Readings
                        LIMSVarConfig.second_imported_tolerance_reading = sheet.Cells(i, 14).Value  # DUT Tolerance
                        LIMSVarConfig.second_imported_device_reading_list.append(LIMSVarConfig.second_imported_dut_reading)
                        LIMSVarConfig.second_imported_reference_reading_list.append(LIMSVarConfig.second_imported_ref_reading)
                        LIMSVarConfig.second_imported_total_error_band_list.append(LIMSVarConfig.second_imported_tolerance_reading)
                        LIMSVarConfig.second_imported_measured_difference_list.append(LIMSVarConfig.second_imported_dut_reading -
                                                                               LIMSVarConfig.second_imported_ref_reading)
                        i += 1
                    else:
                        break

                for i in range(0, len(LIMSVarConfig.second_imported_total_error_band_list)):
                    if -float(LIMSVarConfig.second_imported_total_error_band_list[i]) <= \
                            float(LIMSVarConfig.second_imported_measured_difference_list[i]) <= \
                            float(LIMSVarConfig.second_imported_total_error_band_list[i]):
                        LIMSVarConfig.second_imported_pass_fail_list.append("Pass")
                    else:
                        LIMSVarConfig.second_imported_pass_fail_list.append("Fail")

                # Create Measured Difference List and Pass Fail Criteria

                work_book.Close(True)

                LIMSVarConfig.second_imported_data_checker = int(1)

            else:

                self.__init__()

            certificate_query = tm.askyesno("Complete Calibration?", "Would you like to generate a Certificate of \
Calibration for the device under test? If you would like to select a different file or cancel out of the \
process, click the 'No' button.")

            LIMSVarConfig.imported_data_checker = int(1)

            if certificate_query is True:
                from LIMSCertCreation import AppCalibrationModule
                acm = AppCalibrationModule()
                acm.complete_calibration_process_step_two(data_set_query_window)

            elif certificate_query is False:
                LIMSVarConfig.clear_imported_data_variables()

        else:
            self.__init__()

    # -----------------------------------------------------------------------#

    # Command to extract calibration data from Sonic Nozzle File and store to Arrays for certificate
    def sonic_nozzle_data_import(self):


        sonic_nozzle_search = filedialog.askopenfilename(initialdir="/",
                                                           title="Select Sonic Nozzle Data File",
                                                           filetypes=(("excel files", "*.xls"), ("all files", "*.*")))

        LIMSVarConfig.imported_data_checker = int(0)
        LIMSVarConfig.imported_dut_reading = ""
        LIMSVarConfig.imported_ref_reading = ""
        LIMSVarConfig.imported_device_reading_list = []
        LIMSVarConfig.imported_reference_reading_list = []
        LIMSVarConfig.imported_total_error_band_list = []
        LIMSVarConfig.imported_measured_difference_list = []
        LIMSVarConfig.imported_pass_fail_list = []

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        work_book = excel.Workbooks.Open(sonic_nozzle_search)
        sheet = work_book.Sheets("Sheet1")

        # Extract Data Collected from Sonic Nozzle System
        for i in range(30, 44):
            if sheet.Cells(i, 2).Value is not None:
                LIMSVarConfig.imported_dut_reading = sheet.Cells(i, 2).Value  # Device Under Test Readings
                LIMSVarConfig.imported_device_reading_list.append(LIMSVarConfig.imported_dut_reading)
                LIMSVarConfig.imported_ref_reading = sheet.Cells(i, 3).Value  # Reference Readings
                LIMSVarConfig.imported_reference_reading_list.append(LIMSVarConfig.imported_ref_reading)
                LIMSVarConfig.imported_total_error_band_list.append(((float(LIMSVarConfig.imported_specification_value)
                                                                      / 100)
                                                                     * float(LIMSVarConfig.imported_full_scale)))
                LIMSVarConfig.imported_measured_difference_list.append(LIMSVarConfig.imported_dut_reading -
                                                                       LIMSVarConfig.imported_ref_reading)  # Difference
                i += 1
            else:
                break

        for i in range(0, len(LIMSVarConfig.imported_total_error_band_list)):
            if -float(LIMSVarConfig.imported_total_error_band_list[i]) <= \
                    float(LIMSVarConfig.imported_measured_difference_list[i]) <= \
                    float(LIMSVarConfig.imported_total_error_band_list[i]):
                LIMSVarConfig.imported_pass_fail_list.append("Pass")
            else:
                LIMSVarConfig.imported_pass_fail_list.append("Fail")

        # Create Measured Difference List and Pass Fail Criteria

        work_book.Close(True)

        certificate_query = tm.askyesno("Complete Calibration?", "Would you like to generate a Certificate of \
Calibration for the device under test? If you would like to select a different file or cancel out of the process, \
click the 'No' button.")

        LIMSVarConfig.imported_data_checker = int(1)

        if certificate_query is True:
            from LIMSCertCreation import AppCalibrationModule
            acm = AppCalibrationModule()
            acm.complete_calibration_process_step_two(DataImportSelectionWindow)

        elif certificate_query is False:
            LIMSVarConfig.clear_imported_data_variables()
