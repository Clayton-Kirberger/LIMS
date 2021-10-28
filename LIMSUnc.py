"""
LIMSUnc is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to perform calculations for determining measurement uncertainty and test uncertainty ratios for equipment
used during accredited calibration.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import math
import os

import LIMSVarConfig
import tkinter.messagebox as tm
import win32com.client
from tkinter import *
from tkinter import ttk


class AppUncertainty:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =========================EMU AND TUR CALCULATOR========================= #

    # This command is designed to open up EMU and TUR calculator.
    def uncertainty(self, window):
        self.__init__()
        window.withdraw()

        global unc_option_sel, unc_option_selection

        from LIMSHomeWindow import AppCommonCommands
        acc = AppCommonCommands()

        # ..................Window Characteristics................... #

        unc_option_sel = Toplevel()
        unc_option_sel.title("E.M.U. & T.U.R. Selection")
        unc_option_sel.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 340
        height = 135
        screen_width = unc_option_sel.winfo_screenwidth()
        screen_height = unc_option_sel.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        unc_option_sel.geometry("%dx%d+%d+%d" % (width, height, x, y))
        unc_option_sel.focus_force()
        unc_option_sel.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(unc_option_sel))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(unc_option_sel, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(unc_option_sel)),
                                           ("Logout", lambda: acc.software_signout(unc_option_sel)),
                                           ("Quit", lambda: acc.software_close(unc_option_sel))])

        # ......................Frame Creation...................... #

        uncertainty_selection_frame = LabelFrame(unc_option_sel, text="Please select an option from the drop down \
list", relief=SOLID, bd=1, labelanchor="n")
        uncertainty_selection_frame.grid(row=0, column=0, rowspan=2, padx=10, pady=5)

        # ......................Drop Down Lists..................... #

        if LIMSVarConfig.user_access == 'admin':
            unc_option_selection = ttk.Combobox(uncertainty_selection_frame, values=[" ", "View Scope of Accreditation",
                                                                                     "Calculate EMUs and TURs",
                                                                                     "View Uncertainty Database"])
        else:
            unc_option_selection = ttk.Combobox(uncertainty_selection_frame, values=[" ", "View Scope of Accreditation",
                                                                                     "Calculate EMUs and TURs"])

        acc.always_active_style(unc_option_selection)
        unc_option_selection.configure(state="active", width=25)
        unc_option_selection.focus()
        unc_option_selection.grid(padx=10, pady=5, row=0, column=0)

        # .......................Dummy Labels....................................#

        dummy = Label(uncertainty_selection_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        # ................Button with Functions for this Window...................#

        btn_open_unc_option_selection = ttk.Button(uncertainty_selection_frame, text="Open", width=15,
                                                   command=lambda: self.open_uncertainty_option())
        btn_open_unc_option_selection.bind("<Return>", lambda event: self.open_uncertainty_option())
        btn_open_unc_option_selection.grid(row=0, column=1, padx=5, pady=5)

        btn_backout_unc_option_selection = ttk.Button(uncertainty_selection_frame, text="Back to Main Menu",
                                                      width=20, command=lambda: acc.return_home(unc_option_sel))
        btn_backout_unc_option_selection.bind("<Return>", lambda event: acc.return_home(unc_option_sel))
        btn_backout_unc_option_selection.grid(pady=10, padx=5, row=2, column=0, columnspan=2)

    # ----------------------------------------------------------------------------- #

    # Function to handle uncertainty selection
    def open_uncertainty_option(self):
        if unc_option_selection.get() == "View Scope of Accreditation":
            self.open_scope_of_accreditation()
        elif unc_option_selection.get() == "Calculate EMUs and TURs":
            self.emu_and_tur_calculator(unc_option_sel)
        elif unc_option_selection.get() == "View Uncertainty Database":
            self.open_uncertainty_database()
        else:
            tm.showerror("No E.M.U. & T.U.R. Selection Made", "Please select an action from the drop down provided.")

    # ----------------------------------------------------------------------------- #

    # Function created to allow users to view existing and current Scope of Accreditation
    def open_scope_of_accreditation(self):
        self.__init__()

        current_scope = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Reference Documents\DwyerCertScope-V001.pdf'
        os.startfile(current_scope)

    # ----------------------------------------------------------------------------- #

    # Function created to allow users to calculated EMU and TUR values based on existing uncertainties
    def emu_and_tur_calculator(self, window):
        self.__init__()
        window.withdraw()

        LIMSVarConfig.calculated_emu = ""
        LIMSVarConfig.mass_uncertainty_units = ""
        LIMSVarConfig.calculated_tur = ""

        global new_emu_info, new_emu_tur_frame, metrology_field_selection, parameter_equipment_selection, \
            range_selection, reference_standard, dut_resolution_value, dut_entry_value, dut_tolerance_entry, \
            calculated_emu, uncertainty_units, calculated_tur, btn_update_parameters

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

        # ........................Main Window Properties.........................#

        new_emu_info = Toplevel()
        new_emu_info.title("EMU and TUR Calculator")
        new_emu_info.iconbitmap("\\\\BDC5\\Dwyer Engineering LIMS\\Required Images\\DwyerLogo.ico")
        width = 520
        height = 416
        screen_width = new_emu_info.winfo_screenwidth()
        screen_height = new_emu_info.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        new_emu_info.geometry("%dx%d+%d+%d" % (width, height, x, y))
        new_emu_info.focus_force()
        new_emu_info.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(new_emu_info))

        # .....................Menu Bar Creation..................... #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(new_emu_info, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(new_emu_info)),
                                           ("EMU & TUR Selection", lambda: self.uncertainty(new_emu_info)),
                                           ("Logout", lambda: acc.software_signout(new_emu_info)),
                                           ("Quit", lambda: acc.software_close(new_emu_info))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.emu_and_tur_help())])

        # ..........................Frame Creation...............................#

        new_emu_tur_frame = LabelFrame(new_emu_info, text="EMU & TUR Calculator", relief=SOLID, bd=1, labelanchor="n")
        new_emu_tur_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=5, pady=5)

        # .........................Labels and Entries............................#

        # Metrology Field Selection
        lbl_metrology_field = ttk.Label(new_emu_tur_frame, text="Metrology Discipline:", font=('arial', 12))
        lbl_metrology_field.grid(row=0, padx=5, pady=5)
        metrology_field_selection = ttk.Combobox(new_emu_tur_frame, values=[" ", "Mass and Mass Related"])
        acc.always_active_style(metrology_field_selection)
        metrology_field_selection.configure(state="active", width=27)
        metrology_field_selection.focus()
        metrology_field_selection.grid(row=0, column=1, padx=5, pady=5)

        # Update Parameter/Equipment Dropdown
        btn_update_parameters = ttk.Button(new_emu_tur_frame, text="Update Parameters", width=20,
                                           command=lambda: self.update_emu_tur_parameter())
        btn_update_parameters.bind("<Return>", lambda event: self.update_emu_tur_parameter())
        btn_update_parameters.grid(row=0, column=2, padx=5, pady=5)

        # Parameter/Equipment Selection
        lbl_parameter_equipment = ttk.Label(new_emu_tur_frame, text="Parameter/Equipment:", font=('arial', 12))
        lbl_parameter_equipment.grid(row=1, padx=5, pady=5)
        parameter_equipment_selection = ttk.Combobox(new_emu_tur_frame)
        acc.always_active_style(parameter_equipment_selection)
        parameter_equipment_selection.configure(state="active", width=27)
        parameter_equipment_selection.grid(row=1, column=1, padx=5, pady=5)

        # Update Range Dropdown
        btn_update_range = ttk.Button(new_emu_tur_frame, text="Update Ranges", width=20,
                                      command=lambda: self.update_emu_tur_range())
        btn_update_range.bind("<Return>", lambda event: self.update_emu_tur_range())
        btn_update_range.grid(row=1, column=2, padx=5, pady=5)

        # Range Selection
        lbl_range = ttk.Label(new_emu_tur_frame, text="Range:", font=('arial', 12))
        lbl_range.grid(row=2, padx=5, pady=5)
        range_selection = ttk.Combobox(new_emu_tur_frame, values=[" "])
        acc.always_active_style(range_selection)
        range_selection.configure(state="active", width=27)
        range_selection.grid(row=2, column=1, padx=5, pady=5)

        # Apply Range Information
        btn_apply_range = ttk.Button(new_emu_tur_frame, text="Apply Range", width=20,
                                     command=lambda: self.update_emu_tur_standard())
        btn_apply_range.bind("<Return>", lambda event: self.update_emu_tur_standard())
        btn_apply_range.grid(row=2, column=2, padx=5, pady=5)

        # Reference Standard
        lbl_reference = ttk.Label(new_emu_tur_frame, text="Reference Standard:", font=('arial', 12))
        lbl_reference.grid(row=3, padx=5, pady=5)
        reference_standard = ttk.Label(new_emu_tur_frame, text="", font=('arial', 12))
        reference_standard.grid(row=3, column=1, padx=5, pady=5)

        # DUT Resolution
        lbl_dut_resolution = ttk.Label(new_emu_tur_frame, text="DUT Resolution:", font=('arial', 12))
        lbl_dut_resolution.grid(row=4, padx=5, pady=5)
        dut_resolution_value = ttk.Entry(new_emu_tur_frame)
        dut_resolution_value.config(width=30)
        dut_resolution_value.grid(row=4, column=1, padx=5, pady=5)

        # Nominal Value
        lbl_nominal_value = ttk.Label(new_emu_tur_frame, text="Nominal Value:", font=('arial', 12))
        lbl_nominal_value.grid(row=5, padx=5, pady=5)
        nominal_entry_value = ttk.Entry(new_emu_tur_frame)
        nominal_entry_value.config(width=30)
        nominal_entry_value.grid(row=5, column=1, padx=5, pady=5)

        # DUT Value
        lbl_dut_value = ttk.Label(new_emu_tur_frame, text="DUT Reading:", font=('arial', 12))
        lbl_dut_value.grid(row=6, padx=5, pady=5)
        dut_entry_value = ttk.Entry(new_emu_tur_frame)
        dut_entry_value.config(width=30)
        dut_entry_value.grid(row=6, column=1, padx=5, pady=5)

        # DUT Tolerance
        lbl_dut_tolerance = ttk.Label(new_emu_tur_frame, text="Tolerance Value ("+u"\u00B1"+"):", font=('arial', 12))
        lbl_dut_tolerance.grid(row=7, padx=5, pady=5)
        dut_tolerance_entry = ttk.Entry(new_emu_tur_frame)
        dut_tolerance_entry.config(width=30)
        dut_tolerance_entry.grid(row=7, column=1, padx=5, pady=5)

        # Calculate EMU & TUR
        btn_calculate_emu_tur = ttk.Button(new_emu_tur_frame, text="Calculate", width=20,
                                           command=lambda: self.calculate_emu_tur())
        btn_calculate_emu_tur.bind("<Return>", lambda event: self.calculate_emu_tur())
        btn_calculate_emu_tur.grid(row=7, column=2, padx=5, pady=5)

        # Calculated EMU
        lbl_calc_emu = ttk.Label(new_emu_tur_frame, text="Calculated EMU:", font=('arial', 12))
        lbl_calc_emu.grid(row=8, column=0, padx=5, pady=5)
        calculated_emu = ttk.Label(new_emu_tur_frame, text=LIMSVarConfig.calculated_emu, font=('arial', 12))
        calculated_emu.grid(row=8, column=1, padx=5, pady=5)
        uncertainty_units = ttk.Label(new_emu_tur_frame, text=LIMSVarConfig.mass_uncertainty_units, font=('arial', 12))
        uncertainty_units.grid(row=8, column=2, padx=5, pady=5)

        # Calculated TUR
        lbl_calc_tur = ttk.Label(new_emu_tur_frame, text="Calculated TUR:", font=('arial', 12))
        lbl_calc_tur.grid(row=9, column=0, padx=5, pady=5)
        calculated_tur = ttk.Label(new_emu_tur_frame, text=LIMSVarConfig.calculated_tur, font=('arial', 12))
        calculated_tur.grid(row=9, column=1, padx=5, pady=5)

        # ...............................Buttons..................................#

        btn_backout_emu_tur_calc = ttk.Button(new_emu_info, text="Back", width=20,
                                              command=lambda: self.uncertainty(new_emu_info))
        btn_backout_emu_tur_calc.bind("<Return>", lambda event: self.uncertainty(new_emu_info))
        btn_backout_emu_tur_calc.grid(row=5, column=1, columnspan=2, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to update emu & tur parameter/equipment based on metrology discipline
    def update_emu_tur_parameter(self):
        self.__init__()

        if len(LIMSVarConfig.mass_uncertainty_parameters) < 2:
            btn_update_parameters.config(cursor="watch")
            self.obtain_uncertainty_information()
            btn_update_parameters.config(cursor="arrow")
        else:
            self.__init__()

        if metrology_field_selection.get() == "Mass and Mass Related":
            parameter_equipment_selection.configure(values=LIMSVarConfig.mass_uncertainty_parameters,
                                                    state="active", width=27)
            parameter_equipment_selection.set(LIMSVarConfig.mass_uncertainty_parameters[0])
            parameter_equipment_selection.grid(row=1, column=1, padx=5, pady=5)
        else:
            tm.showerror("No Discipline Selected", "Please select a discipline from the drop down provided and \
then press the 'Update Parameters' button.")

    # ----------------------------------------------------------------------------- #

    # Function created to update emu & tur range based on parameter/equipment selected
    def update_emu_tur_range(self):
        self.__init__()

        LIMSVarConfig.mass_uncertainty_ranges = [" "]
        LIMSVarConfig.mass_uncertainty_units = [" "]
        LIMSVarConfig.mass_uncertainty_range_unit = [" "]

        if parameter_equipment_selection.get() == LIMSVarConfig.mass_uncertainty_parameters[0] or \
                parameter_equipment_selection.get() not in LIMSVarConfig.mass_uncertainty_parameters:
            tm.showerror("No Parameter/Equipment Selected", "Please select a parameter/equipment from the drop down \
provided and then press the 'Update Ranges' button.")
            range_selection.configure(values=LIMSVarConfig.mass_uncertainty_range_unit, state="active", width=27)
            range_selection.grid(row=2, column=1, padx=5, pady=5)
        else:
            for i in range(0, len(local_mass_uncertainty_parameters)):
                if parameter_equipment_selection.get() == local_mass_uncertainty_parameters[i]:
                    LIMSVarConfig.mass_uncertainty_ranges.append(local_mass_uncertainty_ranges[i])
                    LIMSVarConfig.mass_uncertainty_units.append(local_mass_uncertainty_units[i])
                    LIMSVarConfig.mass_uncertainty_range_unit.append(str(local_mass_uncertainty_ranges[i]) +
                                                                     " " + str(local_mass_uncertainty_units[i]))
                    i += 1
                else:
                    i += 1
            range_selection.configure(values=LIMSVarConfig.mass_uncertainty_range_unit, state="active", width=27)
            range_selection.set(LIMSVarConfig.mass_uncertainty_range_unit[0])
            range_selection.grid(row=2, column=1, padx=5, pady=5)
            reference_standard.configure(text="")
            reference_standard.grid(row=3, column=1, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to update reference standard based on range selected
    def update_emu_tur_standard(self):
        self.__init__()

        LIMSVarConfig.mass_reference_standard = " "

        if range_selection.get() == LIMSVarConfig.mass_uncertainty_range_unit[0] or \
                range_selection.get() not in LIMSVarConfig.mass_uncertainty_range_unit:
            tm.showerror("No Range Selected", "Please select a parameter/equipment from the drop down provided and \
then press the 'Apply Range' button.")
            reference_standard.configure(text="")
            reference_standard.grid(row=3, column=1, padx=5, pady=5)
        else:
            for i in range(0, len(local_mass_reference_standard)):
                if local_mass_uncertainty_ranges[i] in range_selection.get():
                    LIMSVarConfig.mass_reference_standard = local_mass_reference_standard[i]
                    i += 1
                else:
                    i += 1
            reference_standard.configure(text=LIMSVarConfig.mass_reference_standard)
            reference_standard.grid(row=3, column=1, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to calculated emu and tur based on user input
    def calculate_emu_tur(self):
        self.__init__()

        s = ttk.Style()
        s.configure('four_to_one.TLabel', background="green")

        s = ttk.Style()
        s.configure('below_four.TLabel', background="yellow")

        s = ttk.Style()
        s.configure('below_two.TLabel', background="red")

        LIMSVarConfig.mass_expanded_uncertainty = " "
        LIMSVarConfig.mass_expanded_uncertainty_floor = " "
        LIMSVarConfig.mass_uncertainty_units = " "

        if dut_resolution_value == "" or dut_entry_value == "" or dut_tolerance_entry == "":
            tm.showerror("Missing Information", "An EMU and TUR could not be calculated. It appears you are missing \
information. Fill out the necessary entry fields an try again.")
        else:
            for i in range(0, len(local_mass_expanded_uncertainty)):
                if local_mass_uncertainty_ranges[i] in range_selection.get():
                    LIMSVarConfig.mass_expanded_uncertainty = float(local_mass_expanded_uncertainty[i])
                    LIMSVarConfig.mass_expanded_uncertainty_floor = float(local_mass_expanded_uncertainty_floor[i])
                    LIMSVarConfig.mass_uncertainty_units = local_mass_uncertainty_units[i]
                    i += 1
                else:
                    i += 1

            LIMSVarConfig.calculated_emu = (float(2.0) * float(dut_resolution_value.get()) * float(1.0/math.sqrt(3.0))) + (float(LIMSVarConfig.mass_expanded_uncertainty)/100.0) * float(dut_entry_value.get()) + float(LIMSVarConfig.mass_expanded_uncertainty_floor)
            LIMSVarConfig.calculated_tur = ((float(dut_entry_value.get()) + float(dut_tolerance_entry.get())) - (float(dut_entry_value.get()) - float(dut_tolerance_entry.get()))) / (2 * LIMSVarConfig.calculated_emu)

            calculated_emu.configure(text=round(LIMSVarConfig.calculated_emu, len(dut_resolution_value.get())-1))
            calculated_emu.grid(row=8, column=1, padx=5, pady=5)

            uncertainty_units.configure(text=LIMSVarConfig.mass_uncertainty_units)
            uncertainty_units.grid(row=8, column=2, padx=5, pady=5)

            if 2.0 < LIMSVarConfig.calculated_tur < 4.0:
                calculated_tur.configure(text=str(round(LIMSVarConfig.calculated_tur, 2)) + ":1",
                                         style='below_four.TLabel')
                calculated_tur.grid(row=9, column=1, padx=5, pady=5)
            elif LIMSVarConfig.calculated_tur < 2.0:
                calculated_tur.configure(text=str(round(LIMSVarConfig.calculated_tur, 2)) + ":1",
                                         style='below_two.TLabel')
                calculated_tur.grid(row=9, column=1, padx=5, pady=5)
            else:
                calculated_tur.configure(text=str(round(LIMSVarConfig.calculated_tur, 2)) + ":1",
                                         style='four_to_one.TLabel')
                calculated_tur.grid(row=9, column=1, padx=5, pady=5)

    # ----------------------------------------------------------------------------- #

    # Function created to write uncertainty database information to arrays
    def obtain_uncertainty_information(self):
        self.__init__()

        global local_mass_uncertainty_parameters, local_mass_uncertainty_ranges, local_mass_uncertainty_units, \
            local_mass_reference_standard, local_mass_expanded_uncertainty, local_mass_expanded_uncertainty_floor

        # Initialize location of mass uncertainty parameters and ranges
        LIMSVarConfig.mass_uncertainty_parameters = [" "]
        local_mass_uncertainty_parameters = [" "]
        local_mass_uncertainty_ranges = [" "]
        local_mass_uncertainty_units = [" "]
        local_mass_expanded_uncertainty = [float()]
        local_mass_expanded_uncertainty_floor = [float()]
        local_mass_reference_standard = [" "]

        excel = win32com.client.dynamic.Dispatch("Excel.Application")
        wkbook = excel.Workbooks.Open(r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Reference Documents\Uncertainty Database.xlsx')
        mass_sheet = wkbook.Sheets("Mass and Mass Related")

        # Log all existing uncertainty parameters to array
        i = 4
        mass_param_cell = mass_sheet.Cells(i, 2)
        for i in range(4, 1000):
            if mass_param_cell.Value is not None:
                local_mass_uncertainty_parameters.append(str(mass_param_cell))
                i += 1
                mass_param_cell = mass_sheet.Cells(i, 2)
            else:
                break

        for entry in local_mass_uncertainty_parameters:
            if entry not in LIMSVarConfig.mass_uncertainty_parameters:
                LIMSVarConfig.mass_uncertainty_parameters.append(entry)

        # Log all existing uncertainty ranges to array
        i = 4
        mass_range_cell = mass_sheet.Cells(i, 3)
        for i in range(4, 1000):
            if mass_range_cell.Value is not None:
                local_mass_uncertainty_ranges.append(str(mass_range_cell))
                i += 1
                mass_range_cell = mass_sheet.Cells(i, 3)
            else:
                break

        # Log all existing uncertainty units to array
        i = 4
        mass_unit_cell = mass_sheet.Cells(i, 5)
        for i in range(4, 1000):
            if mass_unit_cell is not None:
                local_mass_uncertainty_units.append(str(mass_unit_cell))
                i += 1
                mass_unit_cell = mass_sheet.Cells(i, 5)
            else:
                break

        # Log all existing expanded uncertainties to array
        i = 4
        mass_unc_cell = mass_sheet.Cells(i, 6)
        for i in range(4, 1000):
            if mass_unc_cell.Value is not None:
                local_mass_expanded_uncertainty.append(float(mass_unc_cell.Value))
                i += 1
                mass_unc_cell = mass_sheet.Cells(i, 6)
            else:
                break

        # Log all existing expanded uncertainty floors to array
        i = 4
        mass_floor_cell = mass_sheet.Cells(i, 8)
        for i in range(4, 1000):
            if mass_floor_cell.Value is not None:
                local_mass_expanded_uncertainty_floor.append(float(mass_floor_cell.Value))
                i += 1
                mass_floor_cell = mass_sheet.Cells(i, 8)
            else:
                break

        # Log all existing reference standards to array
        i = 4
        mass_std_cell = mass_sheet.Cells(i, 9)
        for i in range(4, 1000):
            if mass_std_cell.Value is not None:
                local_mass_reference_standard.append(str(mass_std_cell))
                i += 1
                mass_std_cell = mass_sheet.Cells(i, 9)
            else:
                break

        wkbook.Close(True)

    # ----------------------------------------------------------------------------- #

    # Function created to allow admin accounts to view uncertainty database
    def open_uncertainty_database(self):
        self.__init__()

        uncertainty_db = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Reference Documents\Uncertainty Database.xlsx'
        os.startfile(uncertainty_db)
