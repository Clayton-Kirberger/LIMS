"""
LIMSAS400Search is a Module/File to be used in conjunction with the main
executable Module/File, Dwyer Engineering LIMS. This module is primarily designed
to log in to AS400 with proper cridentials,

Copyright(c) 2018, Robert Adam Maldonado.
"""

import glob
import os
import os.path
import time

import LIMSVarConfig
import pyautogui


class AppAS400ExecutableHelper:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # -----------------------------------------------------------------------#

    # Open AS400 and Extract Information to .txt File
    # Issue occurs if user password is close to expiring
    # This Needs Work Still
    def as400_executable_helper(self):

        time.sleep(1)
        f = "C:\\Program Files (x86)\IBM\Client Access\Emulator\\pcsfe.exe"
        os.startfile(f)
        time.sleep(5)
        pyautogui.press('left')
        pyautogui.press('right')
        pyautogui.hotkey("Enter")
        time.sleep(5)
        pyautogui.typewrite(LIMSVarConfig.as400_username_helper)
        pyautogui.hotkey("Tab")
        time.sleep(2)
        pyautogui.typewrite(LIMSVarConfig.as400_password_helper)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        if pyautogui.locateOnScreen('Required Images\\AS400 Recover Job.png'):
            pyautogui.typewrite("1")
            pyautogui.hotkey("Enter")
            pyautogui.typewrite("C")
            pyautogui.hotkey("Enter")
            pyautogui.typewrite("D")
            pyautogui.hotkey("Enter")
            pyautogui.typewrite("I")
            pyautogui.hotkey("Enter")
            self.as400_sales_order_search()
        elif pyautogui.locateOnScreen('Required Images\\AS400 Display Messages.png'):
            pyautogui.hotkey("F3")
            time.sleep(2)
            self.as400_sales_order_search()
        elif pyautogui.locateOnScreen(
                'Required Images\\AS400 Authorized Facilities Display.png'):
            pyautogui.hotkey("Enter")
            time.sleep(2)
            self.as400_sales_order_search()
        elif pyautogui.locateOnScreen('Required Images\\AS400 PACS Main Menu 1.png'):
            self.as400_sales_order_search()

    # -----------------------------------------------------------------------#

    # Performs Sales Order Search and Writes Information to Scratch Pad
    def as400_sales_order_search(self):

        pyautogui.typewrite("750")
        pyautogui.hotkey("Enter")
        time.sleep(2)
        pyautogui.typewrite("939")
        pyautogui.hotkey("Enter")
        time.sleep(2)
        pyautogui.typewrite(LIMSVarConfig.as400_sales_order_helper)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        # ..............Write Open Sales Order Display to Scratch Pad.............#
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(5)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        # ..........................Move Back to AS400............................#
        pyautogui.hotkey("Alt", "h")
        time.sleep(2)
        # ...........Search for Sales Order Item and Mark Item for Search.........#
        pyautogui.typewrite('x')
        pyautogui.hotkey("F4")
        # .................Write Order Display to Scratch Pad.....................#
        time.sleep(10)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(3)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("a")
        pyautogui.hotkey("h")
        pyautogui.hotkey("s")
        time.sleep(2)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        pyautogui.hotkey("Alt", "h")
        time.sleep(2)
        # ....................Move to Previous AS400 Screen.......................#
        pyautogui.hotkey("F12")
        time.sleep(2)
        pyautogui.hotkey("F7")
        # ........Write Opens Sales Order Bill-to Display to Scratch Pad..........#
        time.sleep(2)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(2)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("a")
        pyautogui.hotkey("h")
        pyautogui.hotkey("s")
        time.sleep(2)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        pyautogui.hotkey("Alt", "h")
        time.sleep(2)
        pyautogui.hotkey("Shift", "F1")
        # ..................Write Header Notes to Scratch Pad.....................#
        time.sleep(2)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(3)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("a")
        pyautogui.hotkey("h")
        pyautogui.hotkey("s")
        time.sleep(2)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        pyautogui.hotkey("Alt", "h")
        time.sleep(2)
        pyautogui.hotkey("Shift", "F2")
        # ..................Write Item Notes to Scratch Pad.......................#
        time.sleep(2)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(3)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("a")
        pyautogui.hotkey("h")
        pyautogui.hotkey("s")
        time.sleep(2)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Alt", "h")
        time.sleep(2)
        pyautogui.hotkey("Shift", "F3")
        # ..................Write Footer Notes to Scratch Pad.....................#
        time.sleep(2)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(3)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("a")
        pyautogui.hotkey("h")
        pyautogui.hotkey("s")
        time.sleep(2)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(2)
        pyautogui.hotkey("Alt", "h")
        time.sleep(2)
        pyautogui.hotkey("Shift", "F6")
        # .................Write PM Items Notes to Scratch Pad....................#
        time.sleep(2)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("e")
        pyautogui.hotkey("o")
        time.sleep(3)
        pyautogui.hotkey("Alt")
        pyautogui.hotkey("a")
        pyautogui.hotkey("h")
        pyautogui.hotkey("s")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Enter")
        # .................Save Text File as Sales Order Number...................#
        pyautogui.hotkey("Alt", "s")
        time.sleep(3)
        pyautogui.typewrite(LIMSVarConfig.as400_sales_order_helper)
        time.sleep(3)
        pyautogui.hotkey("Alt", "d")
        time.sleep(3)
        pyautogui.typewrite("\\\\BDC5\Dwyer Engineering LIMS\Sales Order Text Files")
        pyautogui.hotkey("Enter")
        time.sleep(3)
        pyautogui.hotkey("Alt", "s")
        time.sleep(3)
        # .............Close Scratch Pad Extension and All of AS400...............#
        pyautogui.hotkey("Alt", "c")
        time.sleep(3)
        pyautogui.hotkey("F3")
        pyautogui.hotkey("F3")
        pyautogui.typewrite("9999")
        pyautogui.hotkey("Enter")
        # .........................Close AS400 Software...........................#
        pyautogui.hotkey("Alt", "f")
        pyautogui.hotkey("x")
        # .......................Close AS400 Emulator Software...................#
        pyautogui.hotkey("Alt", "f")
        pyautogui.hotkey("x")

    # ---------------------------------------------6--------------------------#

    # This function reads through the sales order file generated, saves values
    # to variables created, and loads them into the Calibration_GUI.py executable
    def customer_information_helper(self):

        if os.path.isfile(
                r"\\BDC5\\Dwyer Engineering LIMS\\Sales Order Text Files" + "\\" +
                LIMSVarConfig.external_customer_sales_order_number_helper + ".txt") is True:
            f_sales_order = glob.glob(os.path.join(r"\\BDC5\\Dwyer Engineering LIMS\\Sales Order Text Files",
                                                   LIMSVarConfig.external_customer_sales_order_number_helper + ".*"))[0]
            with open(f_sales_order) as f:
                lines = f.readlines()
            # print "\nSales Order:" + LIMSVarConfig.CustomerSalesOrderNumberHelper + " Pertinent Information\n"
            # print Customer name
            name = lines[52]
            LIMSVarConfig.external_customer_name = name.strip().replace("Name: ", "")
            # print purchase_order Number
            purchase_order = lines[28]
            LIMSVarConfig.external_customer_po = purchase_order.strip().replace("P.O.#: ", "").replace("  ", "")
            # print Customer Address
            address_0 = lines[53]
            LIMSVarConfig.external_customer_address = address_0.strip().replace("Address: ", "")
            address_1 = lines[54]
            LIMSVarConfig.external_customer_address_1 = address_1.strip()
            address_2 = lines[55]
            LIMSVarConfig.external_customer_address_2 = address_2.strip()
            # print Customer city
            city = lines[56]
            LIMSVarConfig.external_customer_city = city.strip().replace("City: ", "")
            # print Customer State, Country, and Zip Code
            state_country_zip = lines[57]
            LIMSVarConfig.external_customer_state_country_zip = (
                state_country_zip.strip().replace("State: ",
                                                  "").replace("Country:",
                                                              " ").replace("  ",
                                                                           "").replace("Zip Code: ",
                                                                                       " ").replace("Zip Code:",
                                                                                                    " ")).lstrip()
            # print rma_value Number if it exists
            rma_value = lines[137]
            LIMSVarConfig.external_customer_rma_number = rma_value.strip().replace("100 ",
                                                                                   "").replace("rma_value",
                                                                                               "").replace(" ",
                                                                                                           "").replace(
                "Y", "").replace("RMA", "")
            if LIMSVarConfig.external_customer_rma_number != "":
                pass
            else:
                LIMSVarConfig.external_customer_rma_number = "-"
            # print Customer Header Notes
            header_notes_0 = lines[87]
            header_notes_1 = lines[88]
            header_notes_2 = lines[89]
            header_notes_3 = lines[90]
            header_notes_4 = lines[91]
            header_notes_5 = lines[92]
            LIMSVarConfig.external_customer_header_notes = (
                header_notes_0.strip().replace("100 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_header_notes_1 = (
                header_notes_1.strip().replace("200 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_header_notes_2 = (
                header_notes_2.strip().replace("300 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_header_notes_3 = (
                header_notes_3.strip().replace("400 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_header_notes_4 = (
                header_notes_4.strip().replace("500 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_header_notes_5 = (
                header_notes_5.strip().replace("600 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            # print Customer Item Notes
            item_notes_0 = lines[112]
            item_notes_1 = lines[113]
            item_notes_2 = lines[114]
            item_notes_3 = lines[115]
            item_notes_4 = lines[116]
            item_notes_5 = lines[117]
            LIMSVarConfig.external_customer_item_notes = (
                item_notes_0.strip().replace("100 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_item_notes_1 = (
                item_notes_1.strip().replace("200 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_item_notes_2 = (
                item_notes_2.strip().replace("300 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_item_notes_3 = (
                item_notes_3.strip().replace("400 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_item_notes_4 = (
                item_notes_4.strip().replace("500 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_item_notes_5 = (
                item_notes_5.strip().replace("600 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            # print Customer Footer Notes
            footer_notes_0 = lines[137]
            footer_notes_1 = lines[138]
            footer_notes_2 = lines[139]
            footer_notes_3 = lines[140]
            footer_notes_4 = lines[141]
            footer_notes_5 = lines[142]
            LIMSVarConfig.external_customer_footer_notes = (
                footer_notes_0.strip().replace("100 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_footer_notes_1 = (
                footer_notes_1.strip().replace("200 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_footer_notes_2 = (
                footer_notes_2.strip().replace("300 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_footer_notes_3 = (
                footer_notes_3.strip().replace("400 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_footer_notes_4 = (
                footer_notes_4.strip().replace("500 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_footer_notes_5 = (
                footer_notes_5.strip().replace("600 ", "").replace(" Y", "")).replace(" N N N", "").rstrip()
            # print Customer PM Item Notes
            pm_item_notes_0 = lines[162]
            pm_item_notes_1 = lines[163]
            pm_item_notes_2 = lines[164]
            pm_item_notes_3 = lines[165]
            pm_item_notes_4 = lines[166]
            pm_item_notes_5 = lines[167]
            LIMSVarConfig.external_customer_pm_item_notes = (
                pm_item_notes_0.replace(" 1 ",
                                        "").strip().replace("100 ",
                                                            "").replace("10 ",
                                                                        "").replace(" Y",
                                                                                    "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_pm_item_notes_1 = (
                pm_item_notes_1.replace(" 2 ",
                                        "").strip().replace("200 ",
                                                            "").replace("20 ",
                                                                        "").replace(" Y",
                                                                                    "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_pm_item_notes_2 = (
                pm_item_notes_2.replace(" 3 ",
                                        "").strip().replace("300 ",
                                                            "").replace("30 ",
                                                                        "").replace(" Y",
                                                                                    "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_pm_item_notes_3 = (
                pm_item_notes_3.replace(" 4 ",
                                        "").strip().replace("400 ",
                                                            "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_pm_item_notes_4 = (
                pm_item_notes_4.replace(" 5 ",
                                        "").strip().replace("500 ",
                                                            "").replace(" Y", "")).replace(" N N N", "").rstrip()
            LIMSVarConfig.external_customer_pm_item_notes_5 = (
                pm_item_notes_5.replace(" 6 ",
                                        "").strip().replace("600 ",
                                                            "").replace(" Y", "")).replace(" N N N", "").rstrip()
        else:
            pass
