"""
LIMSHelpWindows is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
With this module being a standalone and not a part of the Dwyer Engineering LIMS, it can be edited and compiled at will
to add any necessary additions that are deemed to fit the requirements of the laboratory. The primary function of this
module is to be a collection of help windows to be opened throughout the Dwyer Engineering LIMS software application.

Copyright(c) 2018, Robert Adam Maldonado.
"""

from tkinter import messagebox as tm


class AppHelpWindows:

    # ================================METHODS========================================= #

    def __init__(self):
        pass

    # -----------------------------------------------------------------------------#

    # Function to provide user information about each of the modules available on the Home Window
    def home_help(self):
        self.__init__()

        tm.showinfo("'Home' Window Help", "Each of the buttons provided in the Home Window relate to information \
that are a major part in the Engineering Laboratory's day to day operations. \n\n\
* In the 'Laboratory Information' collection, personnel will find tools that include \
'Calculators, Converters & Templates', an 'EMU and TUR Calculator', the company directory to \
officially released 'SPIs, OPs and Lab Procedures', access to the 'Environmental Conditions' for the main lab and \
the flow lab (soon to be prime standards lab), a 'Data Acquisition' module (currently under development) along with \
the essential a 'Certificate of Calibration Creation' module that significantly reduces time to generate certificates \
as well as uses forms that adhere to ISO 17025:2017 requirements.  \n\n\
* In the 'Database Management' collection, personnel will be able to access information regarding \
'Calibration Equipment' used in the laboratory for engineering and calibration purposes, along with existing \
certificates of calibration and logs as they relate to successful calibration and failures in the \
'Certificate Database'. Laboratory personnel will also be able to access RMA Evaluation information in the \
'RMA Database' module. \n\n\
* Lastly, the 'Quality Assurance' collection allows users to be able to provide feedback to the Engineering \
Laboratory Management about the Quality Management System, the Management System itself, and the day to day \
operations of the Engineering Laboratory.")

# ----------------------------------------------------------------------- #

    # Function to provide user information about the current window
    def cct_help(self):
        self.__init__()

        tm.showinfo("'Calculators, Converters & Templates' Help", "All files and programs used by \
laboratory personnel can be found here. Each of the available drop downs provided in \
this window contain files/programs that are the most up-to-date versions. The files are read only \
files so that personnel can save a local copy for their own use. All local files are \
to only be used for that session. Local files must be deleted and the files contained \
in the drop downs are to be treated as master files. \n\n\
To open a calculator, converter or template file/program, simply select one from the drop downs \
provided and click the 'Open' button to the right of the drop down. \n\n\
NOTE: IF THERE ARE SOME FILES/PROGRAMS THAT YOU USE ON A REGULAR BASIS THAT ARE MISSING FROM \
THE DROP DOWNS PROVIDED THAT ARE USEFUL FOR NOT ONLY YOURSELF, BUT ALL TECHNICIANS, \
submit the files to Roger Shumaker or Robert Maldonado for review and to be added to the program.")

# ----------------------------------------------------------------------- #

    # Function to provide user information about the EMU and TUR Calculator Window
    def emu_and_tur_help(self):
        self.__init__()

        tm.showinfo("'EMU and TUR Calculator' Help", "This calculator allows users to accurately calculate Estimated \
Measurement Uncertainty (EMU) values and Test Uncertainty Ratio (TUR) values by using the most recent, approved \
uncertainty budget for the measurement and test equipment (M&TE) used to calibrate or test the device under test. \n\n\
To calculate these values, the user must accurately fill out the entry fields provided (i.e. Nominal Value, \
DUT Reading, DUT Tolerance) as well as select the metrology discipline and uncertainty to be applied to the obtained \
reading. \n\n\
NOTE: TUR is not the same as Test Accuracy Ratio (TAR). TAR only compares the raw accuracy of one device to another. \
TUR is a more accurate comparison between M&TE and a DUT as it takes into account the resolution of the DUT as well \
as the EMU calculated for a device. TUR values can only be calculated if you know the uncertainty contribution \
from your M&TE used for the calibration/testing.")

    # ----------------------------------------------------------------------- #

    # Function to provide user information about the current window
    def environmental_condition_help(self):
        self.__init__()

        tm.showinfo("'Environmental Conditions' Help", "In an effort to maintain laboratory conditions, the \
environmental conditions of both laboratories in which calibration work is done is under regular \
surveillance. \n\n\
Laboratory personnel can obtain the status of the environmental conditions by simply \
clicking either of the 'query' buttons located in this window for Temperature, Pressure, \
Relative Humidity and Dew Point measurements for each of the laboratories respectively. \n\n\
NOTE: Technicians shall review the environmental conditions prior to performing any calibration \
work to ensure the environment is suitable for calibration work.")

    # ----------------------------------------------------------------------- #

    # Function to provide user information about the data acquisition - test setup window
    def data_acq_test_setup_help(self):
        self.__init__()

        tm.showinfo("'Data Acquisition - Test Setup' Help", "The data acquisition function is designed to allow \
users to obtain data on instruments without generating a certificate of calibration. The data obtained is intended \
to be used for test reports and analysis. If the user would like to test an instrument manually at a specific \
location or their work bench, the user should select the 'Manual' test type, apply the settings, select the number \
of models and location in which testing is performed. The user also has the option to perform automated testing \
should they have a large number of samples to be tested or if they simply would like to perform automated \
data acquisition. \n\n\
Depending on the selections made on this screen, the next window will reconfigure to better fit the test type \
selected.")

    # ----------------------------------------------------------------------- #

    # Function to provide user information about the data acquisition - dut model information window
    def data_acq_dut_info_help(self):
        self.__init__()

        tm.showinfo("'Data Acquisition - DUT Model Information' Help", "This window allows the user to input \
information about the device(s) and the respective model(s) to be tested during the data acquisition process. \n\n\
NOTE: The information input in this stage of the data acquisition process will be used for the test report / data \
file generated. This ultimately saves time as the report generated should not need much alteration once it is \
generated.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information about the data acquisition - measurement and test equipment window
    def data_acq_mte_selection_help(self):
        self.__init__()

        tm.showinfo("'Data Acquisition - M&TE Selection & Configuration' Help", "All measurement and test equipment \
that can be used for remote communications for data acquisition are provided in this window. \n\n\
NOTE: Should a device that is used not be listed in this window as an option, please submit an opportunity for \
improvement so that the device can be added as an option. \n\n\
Check the check boxes that are applicable for the devices you are using as your reference/sources. Upon selecting \
the equipment, you will then have the opportunity to configure the communications for each of your devices.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information about the data acquisition - test profile window
    def data_acq_profile_creation_help(self):
        self.__init__()

        tm.showinfo("'Data Acquisition - Test Profile' Help", "This window allows the user to input the \
test parameter information for the type of data acquisition/testing that they would like to perform. Whether the \
testing is manual or automated, the user only has to key in the desired set points, temperature set points \
(if temperature controls are available), soak times, sampling rate and number of loops and the program will create \
the test based on user input. The user can create a new test profile on the fly (and save it for future testing) \
or select an existing test profile (if one exists) for data acquisition. Once the testing is completed, \
a csv file containing all of the data will be generated upon completion along with a formal test report \
detailing the testing performed.")

    # ----------------------------------------------------------------------------- #

    # Function to aid user in understanding how to properly use SO Information Search Module
    def sales_order_import_option_help(self):
        self.__init__()

        tm.showinfo("'Customer Sales Order Information Search' Help", "This window allows users to search \
for a customer sales order and log the information so that the information obtained can be used for \
customer certificates of calibration. \n\n\
The user can also generate a custom certificate of calibration for internal or external customers \
that are not generated in the typical fashion as customers with sales orders or PO's. \n\n\
If a sales order has not had its information logged \
through the automated export function provided, all the user has to do is enter a valid AS400 USERNAME, \
AS400 PASSWORD, and existing sales order in the entry boxes provided and click the 'Search AS400' button. \
When the user presses the 'Search AS400' button, an automated process will begin where the customer information \
is logged to a text file and stored in the database for importing into the certificate of calibration. \
NOTE: ALLOW THE AUTOMATED PROCESS TO FINISH PRIOR TO PERFORMING KEYSTROKES OR MOUSE MOVEMENT. \n\n\
If this automated process has already occurred for a given sales order, there is no need to perform the process \
again. All the user needs to do is click the 'Create Customer Certificate' to begin the certificate of calibration \
process. NOTE: THE INFORMATION MUST BE LOGGED/EXPORTED FIRST PRIOR TO ATTEMPTING GENERATE CERTIFICATE OF \
CALIBRATION. \n\n\
If the user would like to generate a custom certificate of calibration for internal or external customers, \
the user can simply press the 'Generate Custom Certificate' button to begin that process.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use Certificate of Calibration Creation Module
    def external_customer_certificate_information_help(self):
        self.__init__()

        tm.showinfo("'Certificate of Calibration Creation' Help", "This window allows the user to import customer \
sales order information that has been logged through our automation method used on the previous screen. \n\n\
If a sales order HAS NOT had its information logged through the automated export function used on the previous screen, \
the user will need to first perform that step. If a sales order HAS been logged through the automated export function, \
the user only needs to key in a valid/existing sales order number into the appropriate entry field and press the \
'Import Customer Information' button. \n\n\
-When the user presses the 'Import Customer Information' button, all pertinent information documented in the sales \
order generated by Customer Service will be provided in this window. NOTE: SHOULD THE CUSTOMER NOTES PROVIDED EXCEED \
THE SPACE PROVIDED BELOW, THE USER SHOULD ACCESS AS400 TO READ THROUGH THE COMPLETE NOTES PROVIDED.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Details Module
    def device_instrument_information_description_help(self):
        self.__init__()

        tm.showinfo("'Certificate of Calibration Creation - DUT Details' Help", "In this stage of the calibration \
process, the user is able to document all of the information relevant to the calibration performed and the device \
under test. \n\n\
All of the information/fields must be completely filled out prior to moving on to the next stage of the calibration \
process. If there is an empty field, an error will indicate that the remaining information must be filled out. \n\n\
- 'Date Received', 'Date of Calibration' and 'Calibration Due Date' should be keyed in the following format: \
MM/DD/YYYY. \n\n\
- A 'Certificate Number' can be imported from our Certificate of Calibration Database by using either the \
'Import Certificate #' button or the 'Import Cert. & Serial #' button. \n\n\
NOTE: THE 'IMPORT CERT. & SERIAL #' BUTTON SHOULD ONLY BE USED IF THE DEVICE UNDER TEST DOES NOT HAVE A PREVIOUSLY \
ASSIGNED SERIAL NUMBER. \n\n\
- The Condition of the DUT can be selected via the drop down provided. Options include: 'New' (for brand new product), \
'Used' (can be product returned on an RMA) and 'Repaired' (devices that we have provided As Found data for prior to \
sending back to the manufacturer for repair). \n\n\
- The Output Type can be selected via the drop down provided. Options include: 'Single' (for analog devices such as \
pressure gauges, temperature probes, etc.), 'Dual' (for devices with two modes of measurement) and 'Transmitter' \
(for devices with transmitted electrical outputs). \n\n\
- IF YOU IMPORTED A SERIAL NUMBER, the imported serial number will appear as the selected \
Instrument ID Number. OTHERWISE, you can manually key in the serial number for the device (if one exists). \n\n\
- If the customer has provided an Instrument ID Number/Asset Number, you can provide that value in the field \
provided. Otherwise, simply type '-' or 'N/A' in the field provided. \n\n\
- If the device under test has a date code provided (on the unit itself, on the box it came in, or paperwork \
provided), type in that date code in the available entry field. Otherwise, simply type '-' or 'N/A' in the \
field provided. \n\n\
- The Model Number entry field is to include the full model number that is provided on the device.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Details Module
    def device_under_test_specification_details(self):
        self.__init__()

        tm.showinfo("'DUT Calibration Parameters' Help", "Based on the input of the user from the previous \
window, this section of the calibration program customizes its options according to the selected output type of \
the device. \n\n\
All of the information/fields must be completely filled out prior to moving on to the next stage of the calibration \
process. If their is an empty field, an error will indicate that the remaining DUT Calibration Parameters must be \
filled out. \n\n\
- The measurement discipline can be selected via the drop down provided. Disciplines include options such as Flow, \
Humidity, Pressure, Temperature and Velocity. Once a measurement discipline is selected, the user is to press the \
'Update Units' button to populate the 'Units' drop down menu with units of measure corresponding to the discipline \
selected. \n\n\
- A minimum value and maximum value for the device under test must be keyed in to the available entry fields. NOTE: \
BE SURE TO SHOW THE ENTIRE RANGE OF THE DEVICE, EVEN IF YOU ARE ONLY TESTING A SPECIFIC RANGE WITH A \
SPECIFIC ACCURACY. \n\n\
- The measurement resolution displayed by the reference used for calibration can be selected via the provided \
drop down menu. The measurement resolution of the device is to be keyed in via the corresponding entry field. NOTE: \
IF A TRANSMITTER IS BEING CALIBRATED, you can type a '-' or 'N/A' into the DUT Resolution field as the resolution of \
the transmitter is dependent on the multimeter used to monitor the output of the device. \n\n\
- The number of test points for the calibration performed can be selected via the provided drop down list along \
with the direction in which the test points were performed. \n\n\
- Lastly, the accuracy of the device in terms of the range in which you are calibrating should be provided below. \
AT A MINIMUM, the first operator drop down, entry field, and accuracy drop down should be filled out. Should you have \
a compounded accuracy, you can add up to three additional accuracies for the range of the device under test. \n\n\
Once all fields are filled out, you need to press the 'Apply' button to register all selections/choices made prior to \
clicking, the 'Next' button.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Details Module
    def calibration_standard_loading_details(self):
        self.__init__()

        tm.showinfo("'Calibration Reference Standards' Help", "In this stage of the calibration process, the user can \
import calibration standard information for the instrumentation used during calibration of the device under test. \n\n\
NOTE: If calibration was performed on systems such as the Fluke 2468, Sonic Nozzle or Molbox, the user can select to \
import calibration data after loading calibration equipment information. \n\n\
To import information to be used on the certificate of calibration in regards to calibration reference standards \
used during calibration, all the user needs to do is key in a valid/existing asset number for a piece of equipment \
used during calibration into the existing entry fields (the user must fill out asset numbers in order (e.g. 1, 2, 3, \
4, and 5 in that order) and press the 'Load' button. IF A VALID/EXISTING asset number for a device is entered, the \
information corresponding to that reference standard will be loaded. \n\n\
NOTE: The device used to record the Environmental Conditions at the location of the calibration must be recorded. \
This should be the first reference standard you import the information for (i.e. LW-HTP05 for the Flow Lab Omega \
Transmitter, etc.). \n\n\
The personnel filling out this information is fully responsible for verifying the information import is correct in \
relation to the calibration standard and that the calibration standards used are within their calibration interval. \
ALL CERTIFICATES OF CALIBRATION OR WORK DONE USING DEVICES THAT ARE OUTSIDE OF THEIR RESPECTIVE CALIBRATION INTERVALS \
MUST HAVE NOTES EXPLAINING WHY THE NON-CONFORMING WORK IS ALLOWED AND WHO APPROVED THE CALIBRATION STANDARD TO BE \
USED OUTSIDE OF ITS CALIBRATION INTERVAL.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Import System Selection Module
    def data_import_selection_help(self):
        self.__init__()

        tm.showinfo("'Data Import System Selection' Help", "Laboratory personnel can import calibration data obtained \
using one of the provided calibration system options. \n\n\
To import calibration data, simply press one of the buttons corresponding to the system used to obtain data and follow \
the prompts. \n\n\
If the user decides they would rather manually enter in the data, simply go back to the previous page and click the \
'Perform Calibration' button.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use Data Set Query Module
    def data_set_query_help(self):
        self.__init__()

        tm.showinfo("'Data Set Selection' Help", "Laboratory personnel can report calibration data obtained during \
initial calibration, after adjustment, etc. \n\n\
Simply select the data set type from the drop down provided and follow the prompts.")

    # -----------------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use Data Set Query Module
    def dwtester_data_set_query_help(self):
        self.__init__()

        tm.showinfo("'Data Set Selection' Help", "Laboratory personnel can report calibration data obtained during \
initial calibration, after adjustment, etc. \n\n\
For this module, users also need to select a mode of operation in which the calibration was performed. This will \
allow users to correctly select all of the data files obtained during calibration. \n\n\
Simply select the data set type from the drop down provided and follow the prompts.")

    # -----------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Details Module
    def calibration_data_input_help(self):
        self.__init__()

        tm.showinfo("'Calibration Data Input' Help", "Laboratory personnel are to record all data acquired from \
calibration of their device under test. This includes the data obtained from their device under test as well as their \
reference standard used. \n\n\
ONCE all data is obtained/recorded/logged in the respective fields, the user is to click the 'Calculate' button in \
this window for deviations in measurement, pass/fail criteria, and a graph displaying the data to be generated. \n\n\
ONCE the information provided is satisfactory (i.e. error has been calculated, criteria has been listed, graphs have \
been generated), the user is to click the 'Complete' button to provide notes and sign off on the calibration \
performed. \n\n\
WARNING: ONCE YOU CLICK THE 'COMPLETE' BUTTON, YOU CANNOT COME BACK TO CHANGE YOUR DATA. PLEASE ENSURE THAT ALL \
DATA ENTRY IS CORRECT PRIOR TO ADVANCING.")

    # -----------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Details Module
    def technician_notes_and_signature_help(self):
        self.__init__()

        tm.showinfo("'Customer Notes & User Signature' Help", "Laboratory personnel completing a certificate of \
calibration are to input any pertinent information regarding the calibration in the notes section. For example, \
if a flow meter was calibrated with a media such as water or oil, it should be written in the notes section to \
document the calibration media used. \n\n\
Once the user has input their notes, they are to select their name from the drop down provided and click the 'Sign \
and Generate' button. The user will be unable to generate the certificate of calibration until it has been signed \
of the user.")

    # -----------------------------------------------------------------------#

    # Function to aid user in understanding how to properly use DUT Details Module
    def additional_certificate_generation_help(self):
        self.__init__()

        tm.showinfo("'Additional Certificate of Calibration Creation' Help", "Laboratory personnel who have \
multiple units to generate certificates of calibration that are of the exact same model (i.e. with the exact \
same accuracy, range, measurement types, resolution, etc.) can fill out the information provided on this window \
to save time in generating certificates of calibration while also maintaining effectiveness in providing correct \
information. \n\n\
Similar to the calibration process as before, the user is able to import the next available certificate number by \
simply clicking the 'Import Cert.' button provided in the window. Should the user also need a serial number for the \
device (that is, the DUT does not come with a pre-existing serial number), the user should click the 'Import Cert. & \
Serial' button to import both a brand new certificate of calibration number along with a serial number for the DUT \
from the database. \n\n\
NOTE: If the device has a serial number, that information can be keyed into the 'Serial Number' field. If the \
device is coming back on an RMA and has a ID number provided by the customer, that information can be keyed into  the \
'Customer Instrument ID Number' field. \n\n\
ALL INFORMATION MUST BE FILLED OUT prior to advancing to perform the calibration. Should any of the fields not apply \
to the DUT, simply type '-' or 'N/A' into the respective field. \n\n\
If the user has decided that they do not want to perform the calibration, they can simply click the 'Exit' button or \
use any of the functions provided in the toolbar to exit this module of the program. \n\n\
NOTE: ONCE YOU EXIT THIS WINDOW, YOU CANNOT COME BACK. If the user decides they would like to perform the calibration \
after closing out the window, they will need to go through the entire process again for generating a new certificate.")

    # ----------------------------------------------------------------------- #

    # Provides user information about the Calibration Equipment Database window
    def calibration_equipment_database_help(self):
        self.__init__()

        tm.showinfo("'Calibration Equipment Database' Help", "This window allows laboratory personnel to search for \
equipment that is located at this facility (Dwyer Instruments, Inc. - Headquarters located in Michigan City, IN). \n\n\
Laboratory personnel can search equipment information by keying in a specific Asset ID number or specific Serial \
number for a device of interest at the bottom of the screen and clicking the 'search' button. \n\n\
If desired, reports for equipment out of their calibration interval or for equipment that is coming up for \
re-calibration can be generated using the buttons provided at the top of the window for each report respectively.")

    # ----------------------------------------------------------------------- #

    # Provides user information about the CertSearch window
    def certificate_database_help(self):
        self.__init__()

        tm.showinfo("'Certificate of Calibration File Search' Help", "Here, laboratory \
personnel can search for Certificates of Calibration generated through a variety of methods. \n\n\
-To open a directory for certificates for a given year, the user should key in a valid year (i.e. 2018) \
into the box provided and click the 'Search' button inside the box corresponding to that function. \n\n\
--To open a summary sheet/log of certificates of calibration done for a given year, the user should key \
in a valid year (i.e. 2018) into the box provided and click the 'Open' button inside the box corresponding to that \
function. \n\n\
---To open a specific certificate of calibration, the user can simply type in an existing certificate of calibration \
(i.e. 18DWY00-0001) into the box provided and click the 'Open' button inside the box corresponding to that \
function. \n\n\
NOTE: If any values are entered into the box's provided that are invalid (such as years in which calibration \
certificates were not generated at this facility, or certificate numbers that do not exist), the respective \
functions will not open a desired file.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information the RMASearch Window
    def rma_database_help(self):
        self.__init__()

        tm.showinfo("'RMA Database' Help", "To access the laboratories RMA Evaluation summary database/spreadsheet, \
simply enter a valid laboratory RMA Evaluation year into the available entry box, and click the open button. This \
option allows the user to modify existing information, log new information or simply review existing information \
provided. \n\n\
To access specific RMA Evaluations of interest, the user can enter a valid laboratory RMA Evaluation year and RMA \
Evaluation number into the appropriate entry boxes, and click the search button. This will direct the user to the \
specific RMA Evaluation folder (if one exists) containing all of the information provided by personnel who performed \
the RMA Evaluation. \n\
-NOTE: IF NO INFORMATION IS PROVIDED IN THE ENTRY BOXES PRIOR TO CLICKING THE SEARCH BUTTON, the user will be directed \
to the overall laboratories RMA Evaluation Database location. \n\
--NOTE: IF ONLY A VALID RMA EVALUATION YEAR IS PROVIDED IN THE ENTRY BOX, the user will be directed to the \
corresponding RMA Evaluation Database for the provided year. \n\
---NOTE: A VALID RMA EVALUATION YEAR MUST BE PROVIDED IN THE ENTRY BOX IF THE USER IS LOOKING FOR AN RMA \
EVALUATION WITH THE GIVEN RMA EVALUATION NUMBER. IF NO YEAR IS PROVIDED BUT AN RMA NUMBER IS, THE SEARCH FUNCTION \
WILL NOT WORK. \n\n\
The user can also access Service Management using Internet Explorer by clicking on the connect button. \
The user will then be prompted to provide their Service Management Credentials to log in.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information about the Certificate Amendment Window
    def certificate_amendment_help(self):
        self.__init__()

        tm.showinfo("'Certificate Amendment' Help", "In an effort to continuously improve the laboratory management \
system as well as the activities it performs, in the event that a certificate of calibration needs to be amended \
whether it be a mistake acknowledge by a customer or by laboratory personnel, a record of the amendment as well as \
an amended certificate shall be generated. Amendments of any original/existing technical records generated by \
laboratory personnel should be done such that it is compliant with the corporate ISO/IEC 17025:2017 Operating \
Procedure, OP17025-007. \n\n\
To ensure the integrity of the existing record as well as the new record, be sure to fill out each of the required \
sections (i.e. Original Certificate, Amendment Description, Revision Number, Party Requesting Amendment, Reason \
for Amendment, and Effects of Amendment) in their entirety prior to submission and opening of the record to be \
saved as the amendment. After the amendment record has been created and the amended document has been completed, \
submit the amended document for review by Laboratory Management as you would for creating a new certificate.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information for the Complaints Window
    def complaints_help(self):
        self.__init__()

        tm.showinfo("'New Complaint' Help", "In an effort to continuously improve the laboratory management \
system as well as the activities it performs, customers of the laboratory (be it internal or external) are \
engaged in conversations regarding feedback (both positive and negative). Complaints should be filed such that \
they are compliant with the corporate ISO/IEC 17025:2017 Operating Procedure, OP17025-010. \n\n\
Be sure to fill out each of the required sections (i.e. Complaint Location, Complaint Severity, Personnel Involved, \
Complaint Description and Complaint Findings) in their entirety prior to submission. Each Complaint submitted will be \
investigated by Laboratory Management and Laboratory Management will then decide what actions shall be taken in \
response to the complaint filed.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information for the CA Window
    def ca_help(self):
        self.__init__()

        tm.showinfo("'Corrective Action Request Creation' Help", "In an effort to improve the laboratory management \
system as well as the activities it performs, corrective action is taken to eliminate the causes of any existing \
nonconformity, defect, or other undesirable situation in order to prevent recurrence. Corrective Actions should be \
created and implemented such that they are compliant with the corporate ISO/IEC 17025:2017 Operating \
Procedure, OP17025-011. \n\n\
Be sure to fill out each of the required sections (i.e. Investigator Assigned, Associated Risk Level, \
Action Detail, Support Information) in their entirety prior to submission. Each Corrective Action submitted will be \
reviewed for completeness and will need to meet the requirements of being 'fit for use' prior to its implementation.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information for the NC Window
    def nc_help(self):
        self.__init__()

        tm.showinfo("'Nonconforming Work Creation' Help", "In an effort to improve the laboratory management \
system as well as the activities it performs, instances/events of nonconforming work are to be \
identified by the laboratories personnel. Nonconformities can be identified as outlined in the corporate \
ISO/IEC 17025:2017 Operating Procedure, OP17025-011. \n\n\
Be sure to fill out each of the required sections (i.e. Incident Location, Associated Risk Level, Personnel Involved, \
Description and Findings/Evidence) in their entirety prior to submission. Each Nonconformity will be investigated \
by Laboratory Management and Laboratory Management will then decide what actions shall be taken in response to \
the reported incident.")

    # ----------------------------------------------------------------------------- #

    # Function to provide user information for the OFI Window
    def ofi_help(self):
        self.__init__()

        tm.showinfo("'Opportunity for Improvement Creation' Help", "In an effort to improve the laboratory management \
system as well as the activities it performs, opportunities for improvement are to be \
identified by the laboratories personnel. Opportunities for Improvement can be identified as outlined in the corporate \
ISO/IEC 17025:2017 Operating Procedure, OP17025-006. \n\n\
Be sure to fill out each of the sections (i.e. Description, Rationale, Impact and Data) in their entirety prior \
to submission. Each Opportunity for Improvement will be reviewed prior to potential implementation.")

