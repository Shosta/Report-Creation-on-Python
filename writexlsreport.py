"""
Write the eIoT report in an Xls file.\n

The xls file has 4 sheets.\n
1. The first sheet describes all the objects.\n
2. The second sheet decribes the objects that are currently in Evaluation.\n
3. The third sheet decribes the objects that are currently in Qualification.\n
4. The fourth sheet decribes the objects that are currently Audited.
"""

import xlwt

PHASES_DICT = {
    '1': '1. Evaluation des Risques',
    '2': '2. HLRA - Risk Assessments',
    '3': '3. Audits de Securite',
    '4': '4. Tests Manuels',
    '5': '5. Legal and Contracts',
    '6': '6. Delivery and Quality'
}


def write_line_to_report(project_name_string,
                         object_name_string,
                         security_process_phase_string,
                         object_state_string,
                         person_in_charge,
                         result_string,
                         go_no_go_string,
                         excel_sheet,
                         row,
                         column):
    """Write a list of information to a specific row and column to an Excel sheet"""
    excel_sheet.write(row, column, project_name_string)
    excel_sheet.write(row, column+1, object_name_string)
    excel_sheet.write(row, column+2, security_process_phase_string)
    excel_sheet.write(row, column+3, object_state_string)
    excel_sheet.write(row, column+4, person_in_charge)
    excel_sheet.write(row, column+5, result_string)
    excel_sheet.write(row, column+6, go_no_go_string)


def write_header_to_sheet(excel_sheet):
    """Write the header row to a specific Excel Sheet in an Excel file"""

    write_line_to_report('Project Name',
                         'Object Name',
                         'Security Process Phase',
                         'Object State',
                         'Person in Charge',
                         'Result',
                         'Go/No Go',
                         excel_sheet,
                         0,
                         0)


def write_object_status_to_report(object_dictionnary,
                                  excel_sheet,
                                  row,
                                  column):
    """Write an object Dictionnary to a specific row and column into an Excel sheet.
    The Dictionnary key are as following :
    .ProjectName
    .ObjectName
    .SecurityProcessPhase
    .ObjectState
    .PersonInCharge
    .Result
    """
    write_line_to_report(object_dictionnary['ProjectName'],
                         object_dictionnary['ObjectName'],
                         object_dictionnary['SecurityProcessPhase'],
                         object_dictionnary['ObjectState'],
                         object_dictionnary['PersonInCharge'],
                         object_dictionnary['Result'],
                         '',
                         excel_sheet,
                         row,
                         column)


def write_excel_file_from(file_name, objects_dictionnary_array):
    """Write an Excel file (xls) from a array of IoT  Dictionnary objects"""

    # Create the Excel File and the Sheets inside.
    book = xlwt.Workbook(encoding="utf-8")
    excel_sheet_main = book.add_sheet("Suivi Global")
    excel_sheet_suivi_evaluation = book.add_sheet("Evaluation")
    excel_sheet_suivi_qualification = book.add_sheet("Qualification")
    excel_sheet_suivi_delivery = book.add_sheet("Delivery Request")

    # Write Headers for every sheets in Excel file.
    write_header_to_sheet(excel_sheet_main)
    write_header_to_sheet(excel_sheet_suivi_evaluation)
    write_header_to_sheet(excel_sheet_suivi_qualification)
    write_header_to_sheet(excel_sheet_suivi_delivery)

    # Rows for each sheet starts at 1 as row 0 is for the header.
    main_sheet_row = 1
    evaluation_sheet_row = 1
    qualification_sheet_row = 1
    delivery_sheet_row = 1
    # Column starts as 0, as the table starts from the left.
    column = 0

    # Write Content
    for object_dictionnary in objects_dictionnary_array:
        # Write the Object to the main sheet.
        write_object_status_to_report(object_dictionnary,
                                      excel_sheet_main,
                                      main_sheet_row,
                                      column)

        if object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['1']:
            # Write the Object to the "Evaluation des Risques" sheet.
            write_object_status_to_report(object_dictionnary,
                                          excel_sheet_suivi_evaluation,
                                          evaluation_sheet_row,
                                          column)
            evaluation_sheet_row = evaluation_sheet_row + 1

        elif object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['2']:
            # Write the Object to the "HLRA - Risk Assessments" sheet.
            write_object_status_to_report(object_dictionnary,
                                          excel_sheet_suivi_qualification,
                                          qualification_sheet_row,
                                          column)
            qualification_sheet_row = qualification_sheet_row + 1

        elif object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['3']:
            # Write the Object to the "Audits de Securite" sheet.
            write_object_status_to_report(object_dictionnary,
                                          excel_sheet_suivi_delivery,
                                          delivery_sheet_row, column)
            delivery_sheet_row = delivery_sheet_row + 1
        main_sheet_row = main_sheet_row + 1

    book.save(file_name)
