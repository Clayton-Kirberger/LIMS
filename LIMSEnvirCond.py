"""
LIMSEnvirCond is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to perform a query on both the Omega IBHTX-W Transmitter in the Flow Lab along with the Main Lab
for use on calibrations performed in each of the respective laboratories.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import datetime
import subprocess as sub

import LIMSVarConfig
from tkinter import *
from tkinter import messagebox as tm
from tkinter import ttk


class AppEnvironmentalConditions:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # =======================ENVIRONMENTAL CONDITIONS========================= #

    # Environmental Conditions Query Process
    def lab_conditions(self, window):
        window.withdraw()

        global EnvConditions, lbl_main_lab_temperature, lbl_main_lab_pressure, lbl_main_lab_humidity, \
            lbl_main_lab_dewpoint, lbl_flow_lab_temperature, lbl_flow_lab_pressure, lbl_flow_lab_humidity,\
            lbl_flow_lab_dewpoint, btn_main_lab_environment_query, btn_flow_lab_environment_query,\
            lbl_main_lab_condition_timestamp, lbl_flow_lab_condition_timestamp

        from LIMSHomeWindow import AppCommonCommands
        from LIMSHelpWindows import AppHelpWindows
        acc = AppCommonCommands()
        ahw = AppHelpWindows()

    # ......................Main Window Properties............................ #

        EnvConditions = Toplevel()
        EnvConditions.title("Environmental Conditions")
        EnvConditions.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 420
        height = 335
        screen_width = EnvConditions.winfo_screenwidth()
        screen_height = EnvConditions.winfo_screenheight()
        x = (screen_width/2) - (width/2)
        y = (screen_height/2) - (height/2)
        EnvConditions.geometry("%dx%d+%d+%d" % (width, height, x, y))
        EnvConditions.focus_force()
        EnvConditions.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(EnvConditions))

    # .........................Menu Bar Creation.................................. #

        from LIMSHomeWindow import MenuBar
        menubar = MenuBar(EnvConditions, width, height, x, y)
        menubar.add_menu("File", commands=[("Home", lambda: acc.return_home(EnvConditions)),
                                           ("Logout", lambda: acc.software_signout(EnvConditions)),
                                           ("Quit", lambda: acc.software_close(EnvConditions))])
        menubar.add_menu("Help", commands=[("Help", lambda: ahw.environmental_condition_help())])

    # .........................Frame Creation................................. #

        env_conditions_frame = LabelFrame(EnvConditions, text="Main Lab - Environmental Conditions", relief=SOLID,
                                          bd=1, labelanchor="n")
        env_conditions_frame.grid(row=0, column=0, rowspan=3, columnspan=4, padx=10, pady=5)

        env_conditions_frame1 = LabelFrame(EnvConditions, text="Flow Lab - Environmental Conditions", relief=SOLID,
                                           bd=1, labelanchor="n")
        env_conditions_frame1.grid(row=4, column=0, rowspan=3, columnspan=4, padx=10, pady=5)

    # ......................Labels and Entries................................ #
    
        # Temperature Description and Value from Omega IBHTX-W (Main Lab)
        lbl_main_lab_temperature_description = ttk.Label(env_conditions_frame, text="Temperature", font=('arial', 12),
                                                         anchor="n")
        lbl_main_lab_temperature_description.grid(row=1, pady=2, padx=2)
        lbl_main_lab_temperature = ttk.Label(env_conditions_frame, text=LIMSVarConfig.temperature_0, font=('arial', 12),
                                             anchor="n")

        if LIMSVarConfig.temperature_0 == "" or LIMSVarConfig.temperature_0 == " ":
            lbl_main_lab_temperature.grid(row=1, column=1)
            lbl_main_lab_temperature.config(width=12)
        else:
            lbl_main_lab_temperature.grid(row=1, column=1, padx=20)
            lbl_main_lab_temperature.config(width=0)

        # Pressure Description and Value from Omega IBHTX-W (Main Lab)
        lbl_main_lab_pressure_description = ttk.Label(env_conditions_frame, text="Pressure", font=('arial', 12),
                                                      anchor="n")
        lbl_main_lab_pressure_description.grid(row=1, column=2, padx=2)
        lbl_main_lab_pressure = ttk.Label(env_conditions_frame, text=LIMSVarConfig.pressure_0, font=('arial', 12),
                                          anchor="n")

        if LIMSVarConfig.pressure_0 == "" or LIMSVarConfig.pressure_0 == " ":
            lbl_main_lab_pressure.grid(row=1, column=3)
            lbl_main_lab_pressure.config(width=12)
        else:
            lbl_main_lab_pressure.grid(row=1, column=3, padx=20)
            lbl_main_lab_pressure.config(width=0)

        # Humidity Description and Value from Omega IBHTX-W (Main Lab)
        lbl_main_lab_humidity_description = ttk.Label(env_conditions_frame, text="Humidity", font=('arial', 12),
                                                      anchor="n")
        lbl_main_lab_humidity_description.grid(row=2, pady=2, padx=2)
        lbl_main_lab_humidity = ttk.Label(env_conditions_frame, text=LIMSVarConfig.humidity_0, font=('arial', 12),
                                          anchor="n")

        if LIMSVarConfig.humidity_0 == "" or LIMSVarConfig.humidity_0 == " ":
            lbl_main_lab_humidity.grid(row=2, column=1)
            lbl_main_lab_humidity.config(width=12)
        else:
            lbl_main_lab_humidity.grid(row=2, column=1, padx=20)
            lbl_main_lab_humidity.config(width=0)

        # Dewpoint Description and Value from Omega IBHTX-W (Main Lab)
        lbl_main_lab_dewpoint_description = ttk.Label(env_conditions_frame, text="Dewpoint", font=('arial', 12),
                                                      anchor="n")
        lbl_main_lab_dewpoint_description.grid(row=2, column=2, padx=2)
        lbl_main_lab_dewpoint = ttk.Label(env_conditions_frame, text=LIMSVarConfig.dew_point_0, font=('arial', 12),
                                          anchor="n")

        if LIMSVarConfig.dew_point_0 == "" or LIMSVarConfig.humidity_0 == " ":
            lbl_main_lab_dewpoint.grid(row=2, column=3)
            lbl_main_lab_dewpoint.config(width=12)
        else:
            lbl_main_lab_dewpoint.grid(row=2, column=3, padx=20)
            lbl_main_lab_dewpoint.config(width=0)

        # Timestamp of last environmental condition query
        lbl_main_lab_condition_timestamp = ttk.Label(env_conditions_frame, text=LIMSVarConfig.time_date_stamp_helper,
                                                     anchor="n")
        lbl_main_lab_condition_timestamp.grid(row=3, column=0, columnspan=4, pady=2)
    
        # Temperature Description and Value from Omega IBHTX-W (Flow Lab)
        lbl_flow_lab_temperature_description = ttk.Label(env_conditions_frame1, text="Temperature", font=('arial', 12),
                                                         anchor="n")
        lbl_flow_lab_temperature_description.grid(row=5, pady=2, padx=2)
        lbl_flow_lab_temperature = ttk.Label(env_conditions_frame1, text=LIMSVarConfig.temperature_1,
                                             font=('arial', 12), anchor="n")

        if LIMSVarConfig.temperature_1 == "" or LIMSVarConfig.temperature_1 == " ":
            lbl_flow_lab_temperature.grid(row=5, column=1)
            lbl_flow_lab_temperature.config(width=12)
        else:
            lbl_flow_lab_temperature.grid(row=5, column=1, padx=20)
            lbl_flow_lab_temperature.config(width=0)

        # Pressure Description and Value from Omega IBHTX-W (Flow Lab)
        lbl_flow_lab_pressure_description = ttk.Label(env_conditions_frame1, text="Pressure", font=('arial', 12),
                                                      anchor="n")
        lbl_flow_lab_pressure_description.grid(row=5, column=2, padx=2)
        lbl_flow_lab_pressure = ttk.Label(env_conditions_frame1, text=LIMSVarConfig.pressure_1, font=('arial', 12),
                                          anchor="n")

        if LIMSVarConfig.pressure_1 == "" or LIMSVarConfig.pressure_1 == " ":
            lbl_flow_lab_pressure.grid(row=5, column=3)
            lbl_flow_lab_pressure.config(width=12)
        else:
            lbl_flow_lab_pressure.grid(row=5, column=3, padx=20)
            lbl_flow_lab_pressure.config(width=0)

        # Humidity Description and Value from Omega IBHTX-W (Flow Lab)
        lbl_flow_lab_humidity_description = ttk.Label(env_conditions_frame1, text="Humidity", font=('arial', 12),
                                                      anchor="n")
        lbl_flow_lab_humidity_description.grid(row=6, pady=2, padx=2)
        lbl_flow_lab_humidity = ttk.Label(env_conditions_frame1, text=LIMSVarConfig.humidity_1, font=('arial', 12),
                                          anchor="n")

        if LIMSVarConfig.humidity_1 == "" or LIMSVarConfig.humidity_1 == " ":
            lbl_flow_lab_humidity.grid(row=6, column=1)
            lbl_flow_lab_humidity.config(width=12)
        else:
            lbl_flow_lab_humidity.grid(row=6, column=1, padx=20)
            lbl_flow_lab_humidity.config(width=0)

        # Dewpoint Description and Value from Omega IBHTX-W (Flow Lab)
        lbl_flow_lab_dewpoint_description = ttk.Label(env_conditions_frame1, text="Dewpoint", font=('arial', 12),
                                                      anchor="n")
        lbl_flow_lab_dewpoint_description.grid(row=6, column=2, padx=2)
        lbl_flow_lab_dewpoint = ttk.Label(env_conditions_frame1, text=LIMSVarConfig.dew_point_1, font=('arial', 12),
                                          anchor="n")

        if LIMSVarConfig.dew_point_1 == "" or LIMSVarConfig.dew_point_1 == " ":
            lbl_flow_lab_dewpoint.grid(row=6, column=3)
            lbl_flow_lab_dewpoint.config(width=12)
        else:
            lbl_flow_lab_dewpoint.grid(row=6, column=3, padx=20)
            lbl_flow_lab_dewpoint.config(width=0)

        # Timestamp of last environmental condition query
        lbl_flow_lab_condition_timestamp = ttk.Label(env_conditions_frame1, text=LIMSVarConfig.time_date_stamp_helper_1,
                                                     anchor="n")
        lbl_flow_lab_condition_timestamp.grid(row=7, column=0, columnspan=4, pady=2)

    # ..............................Buttons.................................. #

        btn_main_lab_environment_query = ttk.Button(env_conditions_frame, text="Query", width=20,
                                                    command=lambda: self.main_lab_conditions_update())
        btn_main_lab_environment_query.bind("<Return>", lambda event: self.main_lab_conditions_update())
        btn_main_lab_environment_query.grid(row=4, columnspan=4, pady=5)

        btn_flow_lab_environment_query = ttk.Button(env_conditions_frame1, text="Query", width=20,
                                                    command=lambda: self.flow_lab_conditions_update())
        btn_flow_lab_environment_query.bind("<Return>", lambda event: self.flow_lab_conditions_update())
        btn_flow_lab_environment_query.grid(row=8, columnspan=4, pady=5)

        btn_backout_environmental_condition = ttk.Button(EnvConditions, text="Back", width=20,
                                                         command=lambda: acc.return_home(EnvConditions))
        btn_backout_environmental_condition.bind("<Return>", lambda event: acc.return_home(EnvConditions))
        btn_backout_environmental_condition.grid(pady=5, row=9, columnspan=4)

    # ----------------------------------------------------------------------- #

    # Acquire Envinromental Conditions from Omega iBTHX-W in Main Lab
    def main_lab_environmental_conditions_query(self):
        create_no_window = 0x08000000
        # Obtain Ambient Temperature in F
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRTF\r" -C 1 172.20.5.102:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.temperature_0 = (stdout.decode().lstrip('00') + " F")
        # Obtain Atmospheric Pressure in inHg
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRHi\r" -C 1 172.20.5.102:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.pressure_0 = (stdout.decode().lstrip('00') + " inHg")
        # Obtain Ambient Relative Humidity in %RH
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRH2\r" -C 1 172.20.5.102:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.humidity_0 = (stdout.decode().lstrip('00') + " %RH")
        # Obtain Ambient Dewpoint Pressure in F
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRDF2\r" -C 1 172.20.5.102:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.dew_point_0 = (stdout.decode().lstrip('00') + " F")

    # ----------------------------------------------------------------------- #

    # Command to Update ML Variables for Environmental Conditions Window
    def main_lab_conditions_update(self):
        btn_main_lab_environment_query.config(cursor="watch")
        main_lab_conditions = AppEnvironmentalConditions()
        main_lab_conditions.main_lab_environmental_conditions_query()
        #main_lab_conditions.main_lab_conditions_check()
        lbl_main_lab_temperature.config(text=LIMSVarConfig.temperature_0, width=0)
        lbl_main_lab_temperature.grid(padx=20)
        lbl_main_lab_pressure.config(text=LIMSVarConfig.pressure_0, width=0)
        lbl_main_lab_pressure.grid(padx=20)
        lbl_main_lab_humidity.config(text=LIMSVarConfig.humidity_0, width=0)
        lbl_main_lab_humidity.grid(padx=20)
        lbl_main_lab_dewpoint.config(text=LIMSVarConfig.dew_point_0, width=0)
        lbl_main_lab_dewpoint.grid(padx=20)

        # Date and time capture for last query
        now = datetime.datetime.now()
        todays_date = now.strftime("%m-%d-%y").replace("-", "/")
        todays_time = now.strftime("%I-%M-%S %p").replace("-", ":")
        LIMSVarConfig.time_date_stamp_helper = "Last Query: " + todays_date + " - " + todays_time

        lbl_main_lab_condition_timestamp.config(text=LIMSVarConfig.time_date_stamp_helper)
        btn_main_lab_environment_query.config(cursor="arrow")

    # ----------------------------------------------------------------------- #

    # Command to Check if Temperature and Humidity are within Main Lab Environmental Condition Specifications
    def main_lab_conditions_check(self):

        main_lab_temperature = LIMSVarConfig.temperature_0.strip(" F")
        main_lab_humidity = LIMSVarConfig.humidity_0.strip(" %RH")

        if float(62.6) < float(main_lab_temperature) < float(84.2):
            pass
        else:
            lbl_main_lab_temperature.config(background="red")
            tm.showerror("Environmental Conditions",
                         "Temperature not within range of specified environmental condition temperature range.")

        if float(10.0) < float(main_lab_humidity) < float(60.0):
            pass
        else:
            lbl_main_lab_humidity.config(background="red")
            tm.showerror("Environmental Conditions",
                         "Relative Humidity not within range of specified environmental condition humidity range.")
        
    # ----------------------------------------------------------------------- #

    # Acquire Envinromental Conditions from Omega iBTHX-W in Flow Lab
    def flow_lab_conditions_query(self):
        create_no_window = 0x08000000
        # Obtain Ambient Temperature in F
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRTF\r" -C 1 172.20.5.176:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.temperature_1 = (stdout.decode().lstrip('00') + " F")
        # Obtain Atmospheric Pressure in inHg
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRHi\r" -C 1 172.20.5.176:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.pressure_1 = (stdout.decode().lstrip('00') + " inHg")
        # Obtain Ambient Relative Humidity in %RH
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRH2\r" -C 1 172.20.5.176:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.humidity_1 = (stdout.decode().lstrip('00') + " %RH")
        # Obtain Ambient Dewpoint Pressure in F
        args = '//BDC5/Dwyer Engineering LIMS/Required Files/httpget -r -q -S "*SRDF2\r" -C 1 172.20.5.176:2000\n'
        p = sub.Popen(args, stdin=sub.PIPE, stdout=sub.PIPE, stderr=sub.PIPE, creationflags=create_no_window)
        stdout, stderr = p.communicate()
        LIMSVarConfig.dew_point_1 = (stdout.decode().lstrip('00') + " F")

    # ----------------------------------------------------------------------- #

    # Command to Update FL Variables for Environmental Conditions Window
    def flow_lab_conditions_update(self):
        btn_flow_lab_environment_query.config(cursor="watch")
        flow_lab_conditions = AppEnvironmentalConditions()
        flow_lab_conditions.flow_lab_conditions_query()
        #flow_lab_conditions.flow_lab_conditions_check()
        lbl_flow_lab_temperature.config(text=LIMSVarConfig.temperature_1, width=0)
        lbl_flow_lab_temperature.grid(padx=20)
        lbl_flow_lab_pressure.config(text=LIMSVarConfig.pressure_1, width=0)
        lbl_flow_lab_pressure.grid(padx=20)
        lbl_flow_lab_humidity.config(text=LIMSVarConfig.humidity_1, width=0)
        lbl_flow_lab_humidity.grid(padx=20)
        lbl_flow_lab_dewpoint.config(text=LIMSVarConfig.dew_point_1, width=0)
        lbl_flow_lab_dewpoint.grid(padx=20)

        # Date and time capture for last query
        now = datetime.datetime.now()
        todays_date = now.strftime("%m-%d-%y").replace("-", "/")
        todays_time = now.strftime("%I-%M-%S %p").replace("-", ":")
        LIMSVarConfig.time_date_stamp_helper_1 = "Last Query: " + todays_date + " - " + todays_time

        lbl_flow_lab_condition_timestamp.config(text=LIMSVarConfig.time_date_stamp_helper_1)
        btn_flow_lab_environment_query.config(cursor="arrow")

    # ----------------------------------------------------------------------- #

    # Command to Check if Temperature and Humidity are within Flow Lab Environmental Condition Specifications
    def flow_lab_conditions_check(self):

        flow_lab_temperature = LIMSVarConfig.temperature_1.strip(" F")
        flow_lab_humidity = LIMSVarConfig.humidity_1.strip(" %RH")

        if float(66.2) < float(flow_lab_temperature) < float(73.4):
            pass
        else:
            lbl_flow_lab_temperature.config(bg="red")
            tm.showerror("Environmental Conditions",
                         "Temperature not within range of specified environmental condition temperature range.")

        if float(25.0) < float(flow_lab_humidity) < float(65.0):
            pass
        else:
            lbl_flow_lab_humidity.config(bg="red")
            tm.showerror("Environmental Conditions",
                         "Relative Humidity not within range of specified environmental condition humidity range.")
