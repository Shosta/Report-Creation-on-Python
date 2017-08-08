"""
Write the eIoT report in an Xlsx file.\n

The xls file has 4 sheets.

1. The first sheet describes all the objects.\n
2. The second sheet decribes the objects that are currently in Evaluation.\n
3. The third sheet decribes the objects that are currently in Qualification.\n
4. The fourth sheet decribes the objects that are currently Audited.
"""

import xlsxwriter
import variables

PHASES_DICT = {
    'RiskEvaluationKey': '1. Evaluation des Risques',
    'RiskAssessmentKey': '2. HLRA - Risk Assessments',
    'SecurityAuditKey': '3. Audits de Securite',
    'ManualTestKey': '4. Tests Manuels',
    'LegalKey': '5. Legal and Contracts',
    'DeliveryAndQualityKey': '6. Delivery and Quality'
}


def write_bold_line_to_report(project_name_string,
                              object_name_string,
                              security_process_phase_string,
                              object_state_string,
                              person_in_charge,
                              result_string,
                              go_no_go_string,
                              excel_sheet,
                              row,
                              cell_format):
    """Write a list of information to a specific row and column to an Excel sheet as bold."""
    column = 0

    excel_sheet.write(row, column, project_name_string, cell_format)
    excel_sheet.write(row, column+1, object_name_string, cell_format)
    excel_sheet.write(row, column+2, security_process_phase_string, cell_format)
    excel_sheet.write(row, column+3, object_state_string, cell_format)
    excel_sheet.write(row, column+4, person_in_charge, cell_format)
    excel_sheet.write(row, column+5, result_string, cell_format)
    excel_sheet.write(row, column+6, go_no_go_string, cell_format)


def write_line_to_report(project_name_string,
                         object_name_string,
                         security_process_phase_string,
                         object_state_string,
                         person_in_charge,
                         result_string,
                         go_no_go_string,
                         excel_workbook,
                         excel_sheet,
                         row):
    """Write a list of information to a specific row and column to an Excel sheet."""
    column = 0

    red_bold = excel_workbook.add_format({'bold': True, 'font_color': 'red', 'border': 1})
    orange_bold = excel_workbook.add_format({'bold': True, 'font_color': 'orange', 'border': 1})
    green_bold = excel_workbook.add_format({'bold': True, 'font_color': 'green', 'border': 1})

    border = excel_workbook.add_format({'border': 1})

    excel_sheet.write(row, column, project_name_string, border)
    excel_sheet.write(row, column+1, object_name_string, border)
    excel_sheet.write(row, column+2, security_process_phase_string, border)

    # Add format to the Object State row => Red, Amber, Green.
    if object_state_string == variables.NOT_STARTED_PHASE or object_state_string == variables.NOT_EVALUATED_PHASE:
        excel_sheet.write(row, column+3, object_state_string, red_bold)
    elif object_state_string == variables.IN_PROGRESS_PHASE:
        excel_sheet.write(row, column+3, object_state_string, orange_bold)
    elif object_state_string == variables.DONE_PHASE:
        excel_sheet.write(row, column+3, object_state_string, green_bold)

    excel_sheet.write(row, column+4, person_in_charge, border)

    # Add format to the Result row => Red, Amber, Green.
    if result_string.upper().find('HIGH') != -1 and result_string != '':
        excel_sheet.write(row, column+5, result_string, red_bold)
    elif result_string.upper().find('LIGHT') != -1 and result_string != '':
        excel_sheet.write(row, column+5, result_string, orange_bold)
    else:
        excel_sheet.write(row, column+5, result_string, green_bold)

    # Add format to the Go/No Go State row => Red, Amber, Green.
    if go_no_go_string == 'Go':
        excel_sheet.write(row, column+6, go_no_go_string, green_bold)
    elif go_no_go_string == 'No Go':
        excel_sheet.write(row, column+6, go_no_go_string, red_bold)
    else:
        excel_sheet.write(row, column+6, '', border)


def write_header_to_sheet(excel_workbook, excel_sheet):
    """Write the header row to a specific Excel Sheet in an Excel file"""
    header_format = excel_workbook.add_format({'bold': True, 'border': 6})

    write_bold_line_to_report('Project Name',
                              'Object Name',
                              'Security Process Phase',
                              'Object State',
                              'Person in Charge',
                              'Result',
                              'Go/No Go',
                              excel_sheet,
                              0,
                              header_format)


def write_object_status_to_report(object_dictionnary,
                                  excel_workbook,
                                  excel_sheet,
                                  row):
    """Write an object Dictionnary to a specific row and column into an Excel sheet.
    The Dictionnary key are as following :
    .ProjectName
    .ObjectName
    .SecurityProcessPhase
    .ObjectState
    .PersonInCharge
    .Result
    .GoNoGo"""

    write_line_to_report(object_dictionnary['ProjectName'],
                         object_dictionnary['ObjectName'],
                         object_dictionnary['SecurityProcessPhase'],
                         object_dictionnary['ObjectState'],
                         object_dictionnary['PersonInCharge'],
                         object_dictionnary['Result'],
                         object_dictionnary['GoNoGo'],
                         excel_workbook,
                         excel_sheet,
                         row)


def count_done_objects(objects_dictionnary_array):
    ''' Return the number of objects that are in the 'Done' security phase.

    Params:
    objects_dictionnary_array: an Array of Dictionary that describes the IoT object.

    Return:
    The number of objects that are in the 'Done' security phase.'''
    done_objects_count = 0
    for iot_object in objects_dictionnary_array:
        if iot_object['ObjectState'] == variables.DONE_PHASE:
            done_objects_count = done_objects_count + 1

    return done_objects_count


def count_in_progress_objects(objects_dictionnary_array):
    ''' Return the number of objects that are in the 'In Progress' security phase.

    Params:
    objects_dictionnary_array: an Array of Dictionary that describes the IoT object.

    Return:
    The number of objects that are in the 'In Progress' security phase.'''
    in_progress_objects_count = 0
    for iot_object in objects_dictionnary_array:
        if iot_object['ObjectState'] == variables.IN_PROGRESS_PHASE:
            in_progress_objects_count = in_progress_objects_count + 1

    return in_progress_objects_count


def count_not_started_objects(objects_dictionnary_array):
    ''' Return the number of objects that are in the 'Not Started' security phase.

    Params:
    objects_dictionnary_array: an Array of Dictionary that describes the IoT object.

    Return:
    The number of objects that are in the 'Not Started' security phase.'''
    not_started_objects_count = 0
    for iot_object in objects_dictionnary_array:
        if iot_object['ObjectState'] == variables.NOT_STARTED_PHASE:
            not_started_objects_count = not_started_objects_count + 1

    return not_started_objects_count


def count_not_evaluated_objects(objects_dictionnary_array):
    ''' Return the number of objects that are in the 'Not Evaluated' security phase.

    Params:
    objects_dictionnary_array: an Array of Dictionary that describes the IoT object.

    Return:
    The number of objects that are in the 'Not Evaluated' security phase.'''
    not_evaluated_objects_count = 0
    for iot_object in objects_dictionnary_array:
        if iot_object['ObjectState'] == variables.NOT_EVALUATED_PHASE:
            not_evaluated_objects_count = not_evaluated_objects_count + 1

    return not_evaluated_objects_count


def write_excel_file_from(file_name, objects_dictionnary_array):
    """Write an Excel file (xlsx) from an array of IoT  Dictionnary objects"""

    # Create a workbook and add the worksheets.
    workbook = xlsxwriter.Workbook(file_name)

    # Add the worksheets to the workbook.
    excel_sheet_main = workbook.add_worksheet("Suivi Global")
    excel_sheet_suivi_evaluation = workbook.add_worksheet("Evaluation")
    excel_sheet_suivi_qualification = workbook.add_worksheet("Qualification")
    excel_sheet_suivi_delivery = workbook.add_worksheet("Delivery Request")

    # Write Headers for every sheets in Excel file.
    write_header_to_sheet(workbook, excel_sheet_main)
    write_header_to_sheet(workbook, excel_sheet_suivi_evaluation)
    write_header_to_sheet(workbook, excel_sheet_suivi_qualification)
    write_header_to_sheet(workbook, excel_sheet_suivi_delivery)

    # Rows for each sheet starts at 1 as row 0 is for the header.
    main_sheet_row = 1
    evaluation_sheet_row = 1
    qualification_sheet_row = 1
    delivery_sheet_row = 1

    # Write Content
    for object_dictionnary in objects_dictionnary_array:
        # Write all objects to the Main sheet.
        write_object_status_to_report(object_dictionnary,
                                      workbook,
                                      excel_sheet_main,
                                      main_sheet_row)

        main_sheet_row = main_sheet_row + 1

        if object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['RiskEvaluationKey']:
            # Write the Object to the "Evaluation des Risques" sheet.
            write_object_status_to_report(object_dictionnary,
                                          workbook,
                                          excel_sheet_suivi_evaluation,
                                          evaluation_sheet_row)

            evaluation_sheet_row = evaluation_sheet_row + 1

        elif object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['RiskAssessmentKey']:
            # Write the Object to the "HLRA - Risk Assessments" sheet.
            write_object_status_to_report(object_dictionnary,
                                          workbook,
                                          excel_sheet_suivi_qualification,
                                          qualification_sheet_row)
            qualification_sheet_row = qualification_sheet_row + 1

        elif object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['SecurityAuditKey']:
            # Write the Object to the "Audits de Securite" sheet.
            write_object_status_to_report(object_dictionnary,
                                          workbook,
                                          excel_sheet_suivi_delivery,
                                          delivery_sheet_row)
            delivery_sheet_row = delivery_sheet_row + 1

    # Write the number of objects that fulfill the security phases.
    underline = workbook.add_format({'underline': 1, 'font_color': 'black', 'border': 1})
    bold = workbook.add_format({'bold': True, 'font_color': 'grey', 'border': 1})

    excel_sheet_main.write_rich_string(
        'I2',
        underline, 'Nombre d\'objets à l\'état \'Done\':',
        bold, ' ' + str(count_done_objects(objects_dictionnary_array)))

    excel_sheet_main.write_rich_string(
        'I4',
        underline, 'Nombre d\'objets à l\'état \'In Progress\':',
        bold, ' ' + str(count_in_progress_objects(objects_dictionnary_array)))

    excel_sheet_main.write_rich_string(
        'I6',
        underline, 'Nombre d\'objets à l\'état \'Not Started\':',
        bold, ' ' + str(count_not_started_objects(objects_dictionnary_array)))

    excel_sheet_main.write_rich_string(
        'I8',
        underline, 'Nombre d\'objets à l\'état \'Not Evaluated\':',
        bold, ' ' + str(count_not_evaluated_objects(objects_dictionnary_array)))

    workbook.close()
