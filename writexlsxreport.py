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


def count_objects(objects_dictionnary_array):
    ''' Return the number of objects that are in the 'Done',
    'In Progress', 'Not Started' and 'Not Evaluated' security phase.

    Params:
    objects_dictionnary_array: an Array of Dictionary that describes all the IoT objects.

    Return:
    A Dictionnary that describes the number of objects that are in the 'Done',
    'In Progress', 'Not Started' and 'Not Evaluated' security phase.'''
    objects_count_dictionnary = {
        variables.NOT_EVALUATED_PHASE: 0,
        variables.NOT_STARTED_PHASE: 0,
        variables.IN_PROGRESS_PHASE: 0,
        variables.DONE_PHASE: 0
    }
    for iot_object in objects_dictionnary_array:
        if iot_object['ObjectState'] == variables.DONE_PHASE:
            objects_count_dictionnary[variables.DONE_PHASE] = objects_count_dictionnary[variables.DONE_PHASE] + 1
        elif iot_object['ObjectState'] == variables.IN_PROGRESS_PHASE:
            objects_count_dictionnary[variables.IN_PROGRESS_PHASE] = objects_count_dictionnary[variables.IN_PROGRESS_PHASE] + 1
        elif iot_object['ObjectState'] == variables.NOT_STARTED_PHASE:
            objects_count_dictionnary[variables.NOT_STARTED_PHASE] = objects_count_dictionnary[variables.NOT_STARTED_PHASE] + 1
        elif iot_object['ObjectState'] == variables.NOT_EVALUATED_PHASE:
            objects_count_dictionnary[variables.NOT_EVALUATED_PHASE] = objects_count_dictionnary[variables.NOT_EVALUATED_PHASE] + 1

    return objects_count_dictionnary


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


def write_security_manager_information(excel_sheet, excel_workbook):
    '''
    Write the Security Manager information to the Excel file on the right on the Objects table.

    Params:
    excel_sheet: The Excel sheet where we are going to write the Security Manager information.
    excel_workbook: The Workbook from the Excel sheet.
    '''
    # Define the format for the Excel workbook.
    underline = excel_workbook.add_format({'underline': 1, 'font_color': 'black', 'border': 1})
    bold = excel_workbook.add_format({'bold': True, 'font_color': 'black', 'border': 1})

    # Write the Security Manager information to the Excel file on the right on the Objects table.
    excel_sheet.write_rich_string(
        'I1',
        bold, 'Security Manager:',
        underline, ' '  + variables.SECURITY_MANAGER_NAME)
    excel_sheet.write('J1', variables.SECURITY_MANAGER_EMAIL)
    excel_sheet.write('J2', variables.SECURITY_MANAGER_PHONE)


def write_object_counter_to_report(excel_sheet, excel_workbook, objects_dictionnary_array):
    '''Write the number of objects that fulfill the security phases.'''
    # Define the format for the Excel workbook.
    underline = excel_workbook.add_format({'underline': 1, 'font_color': 'black', 'border': 1})
    bold = excel_workbook.add_format({'bold': True, 'font_color': 'black', 'border': 1})

    objects_counter_dictionnary = count_objects(objects_dictionnary_array)

    # Write the counters in the Excel file.
    excel_sheet.write_rich_string(
        'I4',
        underline, 'Nombre d\'objets à l\'état \'Done\':',
        bold, ' ' + str(objects_counter_dictionnary[variables.DONE_PHASE]))

    excel_sheet.write_rich_string(
        'I6',
        underline, 'Nombre d\'objets à l\'état \'In Progress\':',
        bold, ' ' + str(objects_counter_dictionnary[variables.IN_PROGRESS_PHASE]))

    excel_sheet.write_rich_string(
        'I8',
        underline, 'Nombre d\'objets à l\'état \'Not Started\':',
        bold, ' ' + str(objects_counter_dictionnary[variables.NOT_STARTED_PHASE]))

    excel_sheet.write_rich_string(
        'I10',
        underline, 'Nombre d\'objets à l\'état \'Not Evaluated\':',
        bold, ' ' + str(objects_counter_dictionnary[variables.NOT_EVALUATED_PHASE]))


def write_excel_file_from(file_name, objects_dictionnary_array):
    """Write an Excel file (xlsx) from an array of IoT  Dictionnary objects.
    It is going to write 4 sheets.

    1. The first sheet describes all the objects.\n
    2. The second sheet decribes the objects that are currently in Evaluation.\n
    3. The third sheet decribes the objects that are currently in Qualification.\n
    4. The fourth sheet decribes the objects that are currently Audited.

    Params:
    file_name: The Excel file name
    objects_dictionnary_array: The Array that has all the Objects Dictionnary
    """

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

    # Write the Security Manager information to the Excel file on the right on the Objects table.
    write_security_manager_information(excel_sheet_main, workbook)

    # Write the number of objects that fulfill the security phases.
    write_object_counter_to_report(excel_sheet_main, workbook, objects_dictionnary_array)

    workbook.close()
