"""
LIMSCCCreation is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory.  This function contains other
submodules that are useful tools for the personnel generating customized certificates of calibration.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import LIMSVarConfig

from tkinter import *
from tkinter import ttk


class AppCCModule:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # ====================CERTIFICATE OF CALIBRATION CREATION=================#

    def internal_customer_selection(self, window):
        self.__init__()
        window.withdraw()

        global ICInfo, internal_customer_value, internal_customer_displayed_address, \
            internal_customer_displayed_city, internal_customer_displayed_state_city_zip

        from LIMSCertCreation import AppCalibrationModule
        from LIMSHomeWindow import AppCommonCommands
        acm = AppCalibrationModule()
        acc = AppCommonCommands()

        LIMSVarConfig.clear_internal_customer_sales_order_variables()
        LIMSVarConfig.customer_selection_check = int(1)

        # ........................Main Window Properties.........................#

        ICInfo = Toplevel()
        ICInfo.title("Internal Customer Selection")
        ICInfo.iconbitmap("Required Images\\DwyerLogo.ico")
        width = 495
        height = 235
        screen_width = ICInfo.winfo_screenwidth()
        screen_height = ICInfo.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        ICInfo.geometry("%dx%d+%d+%d" % (width, height, x, y))
        ICInfo.focus_force()
        ICInfo.protocol("WM_DELETE_WINDOW", lambda: acc.on_exit(ICInfo))

        # ..........................Frame Creation...............................#

        internal_customer_frame = LabelFrame(ICInfo, text="Customer Information", relief=SOLID, bd=1, labelanchor="n")
        internal_customer_frame.grid(row=0, column=0, rowspan=4, columnspan=4, padx=8, pady=5)

        # .........................Labels and Entries............................#

        # Display Customer name
        lbl_company_name = ttk.Label(internal_customer_frame, text="Customer Name:", font=('arial', 12))
        lbl_company_name.grid(row=0, pady=5, padx=5)
        company_name_value = ttk.Label(internal_customer_frame, text=LIMSVarConfig.internal_customer_name,
                                       font=('arial', 12))
        company_name_value.grid(row=0, column=1, padx=5)

        # Ask for customer location/owner of equipment
        lbl_internal_customer = ttk.Label(internal_customer_frame, text="Location:", font=('arial', 12), anchor="n")
        lbl_internal_customer.grid(row=1, pady=5, padx=5)
        internal_customer_value = ttk.Combobox(internal_customer_frame, values=[" ",
                                                                                "P00 - South Bend",
                                                                                "PSL - Standards Laboratory",
                                                                                "PML - MC Main Laboratory",
                                                                                "P01 - Michigan City",
                                                                                "P02 - Wakarusa", "P09 - EMC-R64",
                                                                                "P11 - W.E. Anderson",
                                                                                "P15 - Wolcott", "P20 - EMC-R54",
                                                                                "P25 - Proximity Controls"])
        acc.always_active_style(internal_customer_value)
        internal_customer_value.config(width=25)
        internal_customer_value.focus()
        internal_customer_value.grid(row=1, column=1, padx=5, sticky="ew")

        # Display internal customer address
        lbl_internal_customer_address = ttk.Label(internal_customer_frame, text="Street Address:", font=('arial', 12))
        lbl_internal_customer_address.grid(row=2, pady=5, padx=5)
        internal_customer_displayed_address = ttk.Label(internal_customer_frame,
                                                        text=LIMSVarConfig.internal_customer_address_displayed,
                                                        font=('arial', 12))
        internal_customer_displayed_address.grid(row=2, column=1, pady=5, padx=5)

        # Display internal customer city
        lbl_internal_customer_city = ttk.Label(internal_customer_frame, text="City:", font=('arial', 12))
        lbl_internal_customer_city.grid(row=3, pady=5, padx=5)
        internal_customer_displayed_city = ttk.Label(internal_customer_frame,
                                                     text=LIMSVarConfig.internal_customer_city_displayed,
                                                     font=('arial', 12))
        internal_customer_displayed_city.grid(row=3, column=1, pady=5, padx=5)

        # Display internal customer State/Country/Zip Code
        lbl_internal_customer_st_ctry_zip = ttk.Label(internal_customer_frame, text="State/Country/Zip Code:",
                                                      font=('arial', 12))
        lbl_internal_customer_st_ctry_zip.grid(row=4, pady=5, padx=5)
        internal_customer_displayed_state_city_zip = ttk.Label(internal_customer_frame,
                                                               text=LIMSVarConfig.internal_customer_state_city_zip_displayed,
                                                               font=('arial', 12))
        internal_customer_displayed_state_city_zip.grid(row=4, column=1, pady=5, padx=5)

        # ...............................Buttons..................................#

        btn_import_internal_customer_information = ttk.Button(internal_customer_frame, text="Load",
                                                              command=lambda: self.load_internal_customer_information_details())
        btn_import_internal_customer_information.bind("<Return>",
                                                      lambda event: self.load_internal_customer_information_details())
        btn_import_internal_customer_information.grid(row=1, column=2, columnspan=2, pady=5, padx=5)
        btn_import_internal_customer_information.config(width=15)

        btn_back_out = ttk.Button(ICInfo, text="Back", command=lambda: acm.customer_calibration_type_selection(ICInfo))
        btn_back_out.bind("<Return>", lambda event: acm.customer_calibration_type_selection(ICInfo))
        btn_back_out.grid(row=10, column=0, columnspan=2, pady=5)
        btn_back_out.config(width=15)

        btn_next = ttk.Button(ICInfo, text="Next",
                              command=lambda: acm.device_instrument_description_information(ICInfo))
        btn_next.bind("<Return>", lambda event: acm.device_instrument_description_information(ICInfo))
        btn_next.grid(row=10, column=2, columnspan=2, pady=5)
        btn_next.config(width=15)

    # -----------------------------------------------------------------------#

    # Load Custom Certificate of Calibration Customer Information
    def load_internal_customer_information_details(self):
        self.__init__()

        LIMSVarConfig.internal_customer_location = internal_customer_value.get()
        if LIMSVarConfig.internal_customer_location == " ":
            LIMSVarConfig.internal_customer_address_displayed = ""
            LIMSVarConfig.internal_customer_city_displayed = ""
            LIMSVarConfig.internal_customer_state_city_zip_displayed = ""
        elif LIMSVarConfig.internal_customer_location == "P00 - South Bend":
            LIMSVarConfig.internal_customer_address_displayed = "6850 Enterprise Dr"
            LIMSVarConfig.internal_customer_city_displayed = "South Bend"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "IN 46628"
        elif LIMSVarConfig.internal_customer_location == "PSL - Standards Laboratory" or \
                LIMSVarConfig.internal_customer_location == "PML - MC Main Laboratory" or \
                LIMSVarConfig.internal_customer_location == "P01 - Michigan City":
            LIMSVarConfig.internal_customer_address_displayed = "102 IN-212"
            LIMSVarConfig.internal_customer_city_displayed = "Michigan City"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "IN 46360"
        elif LIMSVarConfig.internal_customer_location == "P02 - Wakarusa":
            LIMSVarConfig.internal_customer_address_displayed = "55 Ward St"
            LIMSVarConfig.internal_customer_city_displayed = "Wakarusa"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "IN 46573"
        elif LIMSVarConfig.internal_customer_location == "P09 - EMC-R64":
            LIMSVarConfig.internal_customer_address_displayed = "3999 Hupp Rd #R64"
            LIMSVarConfig.internal_customer_city_displayed = "Kingsbury"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "IN 46345"
        elif LIMSVarConfig.internal_customer_location == "P11 - W.E. Anderson":
            LIMSVarConfig.internal_customer_address_displayed = "250 Highgrove Rd"
            LIMSVarConfig.internal_customer_city_displayed = "Grandview"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "MO 64030"
        elif LIMSVarConfig.internal_customer_location == "P15 - Wolcott":
            LIMSVarConfig.internal_customer_address_displayed = "1000 N 900 W"
            LIMSVarConfig.internal_customer_city_displayed = "Wolcott"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "IN 47995"
        elif LIMSVarConfig.internal_customer_location == "P20 - EMC-R54":
            LIMSVarConfig.internal_customer_address_displayed = "3999 Hupp Rd #R54"
            LIMSVarConfig.internal_customer_city_displayed = "Kingsbury"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "IN 46345"
        elif LIMSVarConfig.internal_customer_location == "P25 - Proximity Controls":
            LIMSVarConfig.internal_customer_address_displayed = "1431 MN-210"
            LIMSVarConfig.internal_customer_city_displayed = "Fergus Falls"
            LIMSVarConfig.internal_customer_state_city_zip_displayed = "MN 56537"

        internal_customer_displayed_address.config(text=LIMSVarConfig.internal_customer_address_displayed)
        internal_customer_displayed_city.config(text=LIMSVarConfig.internal_customer_city_displayed)
        internal_customer_displayed_state_city_zip.config(text=LIMSVarConfig.internal_customer_state_city_zip_displayed)

