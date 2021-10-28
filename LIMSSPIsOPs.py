"""
LIMSSPIsOPs is a Module/File to be used in conjunction with the main executable Module/File, Dywer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to open any documentation that is necessary for testing, calibration, etc.

Copyright(c) 2018, Robert Adam Maldonado.
"""

import subprocess as sub


class AppProcedures:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # ======================SPI'S, OP'S & LAB PROCEDURES======================= #

    # Command Designed to Open Dwyer Procedures
    def procedure_documentation(self):
        self.__init__()
        sub.Popen(r'explorer /open, "\\PRODFILE\SPI"' + "'s & OP" + "'s" + "")
