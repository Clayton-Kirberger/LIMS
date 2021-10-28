"""
LIMSVarConfig is a Module/File to be used in conjunction with the main executable Module/File, Dwyer Engineering LIMS.
The primary function of this module is set a default value to all variables in a configuration setting, in which,
these can be imported to other files and edited accordingly.

Copyright(c) 2018, Robert Adam Maldonado.
"""

# ============================Declarations=====================================#

# User Access Rights

user_access = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the EMU & TUR Functions

calculated_emu = ""
calculated_tur = ""
mass_uncertainty_parameters = []
mass_uncertainty_ranges = []
mass_uncertainty_units = []
mass_uncertainty_range_unit = []
mass_expanded_uncertainty = []
mass_expanded_uncertainty_floor = []
mass_reference_standard = []

# Certificate of Calibration Database List

certificate_of_calibration_entry_list_helper = []
certificate_of_calibration_entry_list = []

# Default Time Stamp Value

date_helper = ""
time_date_stamp_helper = ""
time_date_stamp_helper_1 = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Data Acquisition Functions

data_acquisition_test_type = ""
data_acquisition_dut_output_type = ""
data_acquisition_number_of_models = ""
data_acquisition_test_environment = ""
data_acquisition_dut_description_1 = ""
data_acquisition_dut_model_number_1 = ""
data_acquisition_dut_date_code_1 = ""
data_acquisition_mode_of_measure_1 = ""
data_acquisition_dut_units_of_measure_1 = []
data_acquisition_dut_selected_units_1 = ""
mc6_com_port = ""
mc6_baud = ""
mc6_data_bits = ""
mc6_parity = ""
mc6_stop_bit = ""
data_acquisition_profile_name = ""
data_acquisition_test_step = []
data_acquisition_set_point = []
data_acquisition_temperature_set_point = []
data_acquisition_soak_time = []
data_acquisition_sample_rate = []

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Cal Equipment Database Functions

calibration_equipment_reference_array = []
cal_equip_item_no = []
cal_equip_descrip = []
cal_equip_manufact = []
cal_equip_model_no = []
cal_equip_range = []
cal_equip_location = []
cal_equip_asset_no = []
cal_equip_serial_no = []
cal_equip_due_date = []
cal_equip_cal_date = []
cal_equip_cal_interval = []
calibration_equip_description = ""
calibration_equip_manufacturer = ""
calibration_equip_model_number = ""
calibration_equip_range = ""
calibration_equip_location = ""
calibration_equip_asset_num = ""
calibration_equip_serial_num = ""
calibration_equipment_due_date = ""
calibration_equip_cal_date = ""
calibration_equip_cal_interval = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Environmental Conditions Functions

temperature_0 = ""
pressure_0 = ""
humidity_0 = ""
dew_point_0 = ""
temperature_1 = ""
pressure_1 = ""
humidity_1 = ""
dew_point_1 = ""

# Default values assigned to variable associated with Dwyer_Engineering_LIMS.py
# These variables are used for determining user selection and window determination

customer_selection_check = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the AS400 Information Logging Functions

as400_username_helper = ""
as400_password_helper = ""
as400_sales_order_helper = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Certificate of Calibration Functions

external_customer_sales_order_number_helper = ""
external_customer_name = ""
external_customer_address = ""
external_customer_address_1 = ""
external_customer_address_2 = ""
external_customer_city = ""
external_customer_state_country_zip = ""
external_customer_po = ""
external_customer_rma_number = ""
external_customer_header_notes = ""
external_customer_header_notes_1 = ""
external_customer_header_notes_2 = ""
external_customer_header_notes_3 = ""
external_customer_header_notes_4 = ""
external_customer_header_notes_5 = ""
external_customer_item_notes = ""
external_customer_item_notes_1 = ""
external_customer_item_notes_2 = ""
external_customer_item_notes_3 = ""
external_customer_item_notes_4 = ""
external_customer_item_notes_5 = ""
external_customer_footer_notes = ""
external_customer_footer_notes_1 = ""
external_customer_footer_notes_2 = ""
external_customer_footer_notes_3 = ""
external_customer_footer_notes_4 = ""
external_customer_footer_notes_5 = ""
external_customer_pm_item_notes = ""
external_customer_pm_item_notes_1 = ""
external_customer_pm_item_notes_2 = ""
external_customer_pm_item_notes_3 = ""
external_customer_pm_item_notes_4 = ""
external_customer_pm_item_notes_5 = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Custom Certificate of Calibration Functions

internal_customer_name = "Dwyer Instruments, Inc."
internal_customer_location = ""
internal_customer_address_displayed = ""
internal_customer_city_displayed = ""
internal_customer_state_city_zip_displayed = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Equipment Description of the device to be calibrated

certificate_option = ""
calibration_date_helper = ""
calibration_due_date_helper = ""
certificate_of_calibration_number = ""
device_serial_number = ""
certificate_number_checker_criteria = ""
device_date_received_helper = ""
device_model_number_helper = ""
device_identification_number_helper = ""
device_date_code_helper = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Calibration Parameters

calibration_standard_equipment_description = ""
calibration_standard_equipment_description_2 = ""
calibration_standard_equipment_description_3 = ""
calibration_standard_equipment_description_4 = ""
calibration_standard_equipment_description_5 = ""
calibration_standard_equipment_description_6 = ""
calibration_standard_equipment_description_7 = ""
calibration_standard_equipment_description_8 = ""
calibration_standard_serial_number = ""
calibration_standard_serial_number_2 = ""
calibration_standard_serial_number_3 = ""
calibration_standard_serial_number_4 = ""
calibration_standard_serial_number_5 = ""
calibration_standard_serial_number_6 = ""
calibration_standard_serial_number_7 = ""
calibration_standard_serial_number_8 = ""
calibration_standard_equipment_calibration_date = ""
calibration_standard_equipment_calibration_date_2 = ""
calibration_standard_equipment_calibration_date_3 = ""
calibration_standard_equipment_calibration_date_4 = ""
calibration_standard_equipment_calibration_date_5 = ""
calibration_standard_equipment_calibration_date_6 = ""
calibration_standard_equipment_calibration_date_7 = ""
calibration_standard_equipment_calibration_date_8 = ""
calibration_standard_equipment_due_date = ""
calibration_standard_equipment_due_date_2 = ""
calibration_standard_equipment_due_date_3 = ""
calibration_standard_equipment_due_date_4 = ""
calibration_standard_equipment_due_date_5 = ""
calibration_standard_equipment_due_date_6 = ""
calibration_standard_equipment_due_date_7 = ""
calibration_standard_equipment_due_date_8 = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the imported data to be used for generation of calibration certificates

imported_data_checker = int(0)
imported_dut_reading = ""
imported_ref_reading = ""
imported_tolerance_reading = ""
imported_full_scale = ""
imported_specification_value = ""
imported_device_reading_list = []
imported_reference_reading_list = []
imported_total_error_band_list = []
imported_measured_difference_list = []
imported_pass_fail_list = []
second_imported_data_checker = int(0)
second_imported_dut_reading = ""
second_imported_ref_reading = ""
second_imported_tolerance_reading = ""
second_imported_device_reading_list = []
second_imported_reference_reading_list = []
second_imported_total_error_band_list = []
second_imported_measured_difference_list = []
second_imported_pass_fail_list = []


# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Certificate of Calibration Authorization

calibration_time_helper = ""
certificate_technician_notes = ""
certificate_technician_name = ""
certificate_technician_signature = ""
certificate_technician_job_title = ""

# Default values assigned to variables associated with Dwyer Engineering LIMS.py
# These variables are for the Certificate Amendment Functions

cert_amend_original_cert = ""
cert_amend_request_personnel = ""
cert_amend_revision_number = ""
cert_amend_description = ""
cert_amend_reason = ""
cert_amend_effects = ""

# Default values assigned to variables associated with Dwyer Engineering LIMS.py
# These variables are for the Complaints Functions

comp_identifier = ""
comp_location = ""
comp_severity = ""
comp_personnel = ""
comp_description = ""
comp_findings = ""

# Default values assigned to variables associated with Dwyer Engineering LIMS.py
# These variables are for the Corrective Action Functions

ca_identifier = ""
ca_investigator = ""
ca_severity = ""
ca_description = ""
ca_notes = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Nonconforming Work Functions

nc_identifier = ""
nc_location = ""
nc_severity = ""
nc_personnel = ""
nc_description = ""
nc_findings = ""

# Default values assigned to variables associated with Dwyer_Engineering_LIMS.py
# These variables are for the Opportunity For Improvement Functions

ofi_identifier = ""
ofi_description = ""
ofi_rationale = ""
ofi_impact = ""
ofi_data = ""


def clear_calibration_equip_variables():
    global calibration_equip_description, calibration_equip_manufacturer, calibration_equip_location, \
        calibration_equip_cal_interval, calibration_equip_model_number, calibration_equip_range, \
        calibration_equip_cal_date, calibration_equip_asset_num, calibration_equip_serial_num, \
        calibration_equipment_due_date

    calibration_equip_description = ""
    calibration_equip_manufacturer = ""
    calibration_equip_location = ""
    calibration_equip_cal_interval = ""
    calibration_equip_model_number = ""
    calibration_equip_range = ""
    calibration_equip_cal_date = ""
    calibration_equip_asset_num = ""
    calibration_equip_serial_num = ""
    calibration_equipment_due_date = ""


def clear_external_customer_sales_order_variables():
    global external_customer_sales_order_number_helper, external_customer_name, external_customer_address, \
        external_customer_address_1, external_customer_address_2, external_customer_city, \
        external_customer_state_country_zip, external_customer_po, external_customer_rma_number, \
        external_customer_header_notes, external_customer_header_notes_1, external_customer_header_notes_2, \
        external_customer_header_notes_3, external_customer_header_notes_4, external_customer_header_notes_5, \
        external_customer_item_notes, external_customer_item_notes_1, external_customer_item_notes_2, \
        external_customer_item_notes_3, external_customer_item_notes_4, external_customer_item_notes_5, \
        external_customer_footer_notes, external_customer_footer_notes_1, external_customer_footer_notes_2, \
        external_customer_footer_notes_3, external_customer_footer_notes_4, external_customer_footer_notes_5, \
        external_customer_pm_item_notes, external_customer_pm_item_notes_1, external_customer_pm_item_notes_2, \
        external_customer_pm_item_notes_3, external_customer_pm_item_notes_4, external_customer_pm_item_notes_5

    external_customer_sales_order_number_helper = ""
    external_customer_name = ""
    external_customer_address = ""
    external_customer_address_1 = ""
    external_customer_address_2 = ""
    external_customer_city = ""
    external_customer_state_country_zip = ""
    external_customer_po = ""
    external_customer_rma_number = ""
    external_customer_header_notes = ""
    external_customer_header_notes_1 = ""
    external_customer_header_notes_2 = ""
    external_customer_header_notes_3 = ""
    external_customer_header_notes_4 = ""
    external_customer_header_notes_5 = ""
    external_customer_item_notes = ""
    external_customer_item_notes_1 = ""
    external_customer_item_notes_2 = ""
    external_customer_item_notes_3 = ""
    external_customer_item_notes_4 = ""
    external_customer_item_notes_5 = ""
    external_customer_footer_notes = ""
    external_customer_footer_notes_1 = ""
    external_customer_footer_notes_2 = ""
    external_customer_footer_notes_3 = ""
    external_customer_footer_notes_4 = ""
    external_customer_footer_notes_5 = ""
    external_customer_pm_item_notes = ""
    external_customer_pm_item_notes_1 = ""
    external_customer_pm_item_notes_2 = ""
    external_customer_pm_item_notes_3 = ""
    external_customer_pm_item_notes_4 = ""
    external_customer_pm_item_notes_5 = ""


def clear_loaded_calibration_standard_information_variables():
    global calibration_standard_equipment_description, calibration_standard_equipment_description_2, \
        calibration_standard_equipment_description_3, calibration_standard_equipment_description_4, \
        calibration_standard_equipment_description_5, calibration_standard_equipment_description_6, \
        calibration_standard_equipment_description_7, calibration_standard_equipment_description_8, \
        calibration_standard_serial_number, calibration_standard_serial_number_2, \
        calibration_standard_serial_number_3, calibration_standard_serial_number_4, \
        calibration_standard_serial_number_5, calibration_standard_serial_number_6,\
        calibration_standard_serial_number_7, calibration_standard_serial_number_8, \
        calibration_standard_equipment_calibration_date, calibration_standard_equipment_calibration_date_2, \
        calibration_standard_equipment_calibration_date_3, calibration_standard_equipment_calibration_date_4, \
        calibration_standard_equipment_calibration_date_5, calibration_standard_equipment_calibration_date_6, \
        calibration_standard_equipment_calibration_date_7, calibration_standard_equipment_calibration_date_8, \
        calibration_standard_equipment_due_date, calibration_standard_equipment_due_date_2, \
        calibration_standard_equipment_due_date_3, calibration_standard_equipment_due_date_4, \
        calibration_standard_equipment_due_date_5, calibration_standard_equipment_due_date_6, \
        calibration_standard_equipment_due_date_7, calibration_standard_equipment_due_date_8

    calibration_standard_equipment_description = ""
    calibration_standard_equipment_description_2 = ""
    calibration_standard_equipment_description_3 = ""
    calibration_standard_equipment_description_4 = ""
    calibration_standard_equipment_description_5 = ""
    calibration_standard_equipment_description_6 = ""
    calibration_standard_equipment_description_7 = ""
    calibration_standard_equipment_description_8 = ""
    calibration_standard_serial_number = ""
    calibration_standard_serial_number_2 = ""
    calibration_standard_serial_number_3 = ""
    calibration_standard_serial_number_4 = ""
    calibration_standard_serial_number_5 = ""
    calibration_standard_serial_number_6 = ""
    calibration_standard_serial_number_7 = ""
    calibration_standard_serial_number_8 = ""
    calibration_standard_equipment_calibration_date = ""
    calibration_standard_equipment_calibration_date_2 = ""
    calibration_standard_equipment_calibration_date_3 = ""
    calibration_standard_equipment_calibration_date_4 = ""
    calibration_standard_equipment_calibration_date_5 = ""
    calibration_standard_equipment_calibration_date_6 = ""
    calibration_standard_equipment_calibration_date_7 = ""
    calibration_standard_equipment_calibration_date_8 = ""
    calibration_standard_equipment_due_date = ""
    calibration_standard_equipment_due_date_2 = ""
    calibration_standard_equipment_due_date_3 = ""
    calibration_standard_equipment_due_date_4 = ""
    calibration_standard_equipment_due_date_5 = ""
    calibration_standard_equipment_due_date_6 = ""
    calibration_standard_equipment_due_date_7 = ""
    calibration_standard_equipment_due_date_8 = ""


def clear_internal_customer_sales_order_variables():
    global internal_customer_location, internal_customer_address_displayed, internal_customer_city_displayed, \
        internal_customer_state_city_zip_displayed

    internal_customer_location = ""
    internal_customer_address_displayed = ""
    internal_customer_city_displayed = ""
    internal_customer_state_city_zip_displayed = ""


def clear_device_under_test_information_variables():
    global certificate_option, calibration_date_helper, calibration_due_date_helper, \
        certificate_of_calibration_number, device_serial_number, certificate_number_checker_criteria, \
        device_date_received_helper, device_model_number_helper, device_identification_number_helper, \
        device_date_code_helper

    certificate_option = ""
    calibration_date_helper = ""
    calibration_due_date_helper = ""
    certificate_of_calibration_number = ""
    device_serial_number = ""
    certificate_number_checker_criteria = ""
    device_date_received_helper = ""
    device_model_number_helper = ""
    device_identification_number_helper = ""
    device_date_code_helper = ""


def clear_technician_information_variables():
    global calibration_time_helper, certificate_technician_notes, certificate_technician_name, \
        certificate_technician_signature, certificate_technician_job_title

    calibration_time_helper = ""
    certificate_technician_notes = ""
    certificate_technician_name = ""
    certificate_technician_signature = ""
    certificate_technician_job_title = ""


def clear_imported_data_variables():
    global imported_data_checker, imported_dut_reading, imported_ref_reading, imported_full_scale, \
        imported_specification_value, imported_device_reading_list, imported_reference_reading_list, \
        imported_total_error_band_list, imported_measured_difference_list, imported_pass_fail_list, \
        imported_tolerance_reading, second_imported_dut_reading, second_imported_ref_reading, \
        second_imported_tolerance_reading, second_imported_device_reading_list, \
        second_imported_reference_reading_list, second_imported_total_error_band_list, \
        second_imported_measured_difference_list, second_imported_pass_fail_list, second_imported_data_checker

    imported_data_checker = int(0)
    imported_dut_reading = ""
    imported_ref_reading = ""
    imported_tolerance_reading = ""
    imported_full_scale = ""
    imported_specification_value = ""
    imported_device_reading_list = []
    imported_reference_reading_list = []
    imported_total_error_band_list = []
    imported_measured_difference_list = []
    imported_pass_fail_list = []
    second_imported_data_checker = int(0)
    second_imported_dut_reading = ""
    second_imported_ref_reading = ""
    second_imported_tolerance_reading = ""
    second_imported_device_reading_list = []
    second_imported_reference_reading_list = []
    second_imported_total_error_band_list = []
    second_imported_measured_difference_list = []
    second_imported_pass_fail_list = []


def clear_all_certificate_of_calibration_variables():
    global certificate_of_calibration_entry_list, certificate_of_calibration_entry_list_helper

    clear_calibration_equip_variables()
    clear_external_customer_sales_order_variables()
    clear_loaded_calibration_standard_information_variables()
    clear_internal_customer_sales_order_variables()
    clear_device_under_test_information_variables()
    clear_technician_information_variables()
    clear_imported_data_variables()
    certificate_of_calibration_entry_list = []
    certificate_of_calibration_entry_list_helper = []
