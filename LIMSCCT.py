"""
LIMSCCT is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to open all calculators, converters and templates that the laboratory uses on a regular basis that are
essential to its operations in terms of testing performed.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import os
import os.path

from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk


class AppCCT:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =======================CALCULATORS AND CONVERTERS======================= #

    # Add additional spreadsheets and calculators as they are created
    # Command Designed to Open Conversion Calculators

    def conversions_and_calculators(self, window):
        window.withdraw()

        global CalcConvTemp, calc_conv_list, template_list, calculator_converter_type, template_drop_down_type

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

    # ....................Main Window Properties............................. #

        CalcConvTemp = Toplevel()
        CalcConvTemp.title("Calculators, Converters & Templates")
        CalcConvTemp.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 455
        height = 195
        screen_width = CalcConvTemp.winfo_screenwidth()
        screen_height = CalcConvTemp.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        CalcConvTemp.geometry("%dx%d+%d+%d" % (width, height, x, y))
        CalcConvTemp.focus_force()
        CalcConvTemp.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(CalcConvTemp))

    # .........................Menu Bar Creation.................................. #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(CalcConvTemp, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(CalcConvTemp)),
                                           ("Logout", lambda: acc.software_signout(CalcConvTemp)),
                                           ("Quit", lambda: acc.software_close(CalcConvTemp))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.cct_help())])

    # .........................Frame Creation................................ #

        cc_frame = LabelFrame(CalcConvTemp, text="Select Calculator or Converter from Drop Down List",
                              relief=SOLID, bd=1, labelanchor="n")
        cc_frame.grid(row=0, column=0, rowspan=2, columnspan=2, padx=5, pady=5)

        template_frame = LabelFrame(CalcConvTemp, text="Select Template from Drop Down List",
                                    relief=SOLID, bd=1, labelanchor="n")
        template_frame.grid(row=2, column=0, rowspan=2, columnspan=2, padx=5, pady=5)

    # .......................Drop Down Lists................................. #

        template_list = [" ", "Data Acquisition Form", "Master Pressure Gage Accuracy Spreadsheet V.11",
                         "PFA Calculator - Rev. 0", "Rockwell Hardness Testing", "Standard DUT Test Points"]
        template_type_width = len(max(template_list, key=len))

        calculator_converter_type = ttk.Combobox(cc_frame, values=[" ", "RHP Thermistor RTD Calculator",
                                                                   "TAR Calculator", "Master Converter",
                                                                   "Transmitter Output Converter"])
        acc.always_active_style(calculator_converter_type)
        calculator_converter_type.configure(state="active", width=template_type_width)
        calculator_converter_type.focus()
        calculator_converter_type.grid(padx=10, pady=5, row=0, column=0)

        template_drop_down_type = ttk.Combobox(template_frame, values=[" ", "Data Acquisition Form",
                                                                       "Data Template - 34401A",
                                                                       "Data Template - 34450A",
                                                                       "Data Template - MC6, MC2, ATE-2",
                                                                       "Master Pressure Gage Accuracy Spreadsheet V.11",
                                                                       "PFA Calculator - Rev. 0",
                                                                       "Rockwell Hardness Testing",
                                                                       "Standard DUT Test Points"])
        acc.always_active_style(template_drop_down_type)
        template_drop_down_type.configure(state="active", width=template_type_width)
        template_drop_down_type.grid(padx=10, pady=5, row=0, column=0)

    # .......................Dummy Labels.................................... #

        dummy = Label(cc_frame)
        dummy.grid(row=0, column=2)
        dummy.config(width=1)

        dummy1 = Label(template_frame)
        dummy1.grid(row=0, column=2)
        dummy1.config(width=1)

    # ................Button with Functions for this Window................... #

        btn_ccopen = ttk.Button(cc_frame, text="Open", width=15,
                                command=lambda: self.open_calculator_converter())
        btn_ccopen.bind("<Return>", lambda event: self.open_calculator_converter())
        btn_ccopen.grid(row=0, column=1, padx=5, pady=5)

        btn_templateopen = ttk.Button(template_frame, text="Open", width=15,
                                      command=lambda: self.open_template())
        btn_templateopen.bind("<Return>", lambda event: self.open_template())
        btn_templateopen.grid(row=0, column=1, padx=5, pady=5)
    
        btn_backout_cct = ttk.Button(CalcConvTemp, text="Back to Main Menu", width=20,
                                     command=lambda: acc.return_home(CalcConvTemp))
        btn_backout_cct.bind("<Return>", lambda event: acc.return_home(CalcConvTemp))
        btn_backout_cct.grid(pady=10, padx=5, row=4, column=0, columnspan=2)

    # ----------------------------------------------------------------------- #

    # Open Calculator or Converter File Selected from Drop Down
    def open_calculator_converter(self):

        if calculator_converter_type.get() == "RHP Thermistor RTD Calculator":
            calc_conv_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Calculators & Converters\RHP Thermistor RTD Calculator.xls'
            os.startfile(calc_conv_file)
        elif calculator_converter_type.get() == "TAR Calculator":
            calc_conv_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Calculators & Converters\TAR Calculator.xlsx'
            os.startfile(calc_conv_file)
        elif calculator_converter_type.get() == "Master Converter":
            calc_conv_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Calculators & Converters\MasterConverter.exe'
            os.startfile(calc_conv_file)
        elif calculator_converter_type.get() == "Transmitter Output Converter":
            calc_conv_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Calculators & Converters\Transmitter Output Converter.xlsx'
            os.startfile(calc_conv_file)
        else:
            tm.showerror("No File Selected", "Please select a file from the drop down provided!")

    # ----------------------------------------------------------------------- #

    # Open Template File Selected from Drop Down
    def open_template(self):

        if template_drop_down_type.get() == "Data Acquisition Form":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\Data Acquisition Form.xls'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "Data Template - 34401A":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\Data Template - 34401A.xlsx'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "Data Template - 34450A":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\Data Template - 34450A.xlsx'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "Data Template - MC6, MC2, ATE-2":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\Data Template - MC6, MC2, ATE-2.xlsx'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "Master Pressure Gage Accuracy Spreadsheet V.11":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\Master Pressure Gage Accuracy Spreadsheet 32 and 64 Bit  Ver 11.xls'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "PFA Calculator - Rev. 0":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\PFA Calculator - Rev. 0.xlsm'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "Rockwell Hardness Testing":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Templates\Rockwell Hardness Testing - Template.xls'
            os.startfile(template_file)
        elif template_drop_down_type.get() == "Standard DUT Test Points":
            template_file = r'\\BDC5\Dwyer Engineering LIMS\Required Files\LIMS CCT Files\Reference Documents\Standard Test Points.xlsx'
            os.startfile(template_file)
        else:
            tm.showerror("No File Selected", "Please select a file from the drop down provided!")
