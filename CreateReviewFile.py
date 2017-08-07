"""
Write the eIoT report automatically based on the files in the directory tree.
"""

import os.path
import xlwt


PHASES_DICT = {
    '1': '1. Evaluation des Risques',
    '2': '2. HLRA - Risk Assessments',
    '3': '3. Audits de Securite',
    '4': '4. Tests Manuels',
    '5': '5. Legal and Contracts',
    '6': '6. Delivery and Quality'
}

NOT_STARTED_PHASE = 'Not Started'
IN_PROGRESS_PHASE = 'In Progress'
DONE_PHASE = 'Done'
SEPARATOR = ' - '
CURRENT_DIR = '.'

NOT_STARTED_PROJECTS_COUNTER = 0
IN_PROGRESS_PROJECTS_COUNTER = 0
DONE_PROJECTS_COUNTER = 0

def substring_from(string, separator, index):
    """Retrieve the string from a larger string splitted regarding the separator
    and returning the indexed one.
    """

    liste = string.split(separator)
    object_name = ""
    try:
        object_name = liste[index]
    except IndexError:
        if index-1 >= 0:
            object_name = substring_from(string, separator, index-1)


    return object_name

def phase_name_from(directory_path):
    """Get the available phases names from the directory tree."""

    #Linux
    directory_separator = '/'
    #Windows
    directory_separator = '\\'

    phase_name = ''
    try:
        phase_name = substring_from(directory_path,
                                    directory_separator,
                                    1)
    except IndexError:
        pass
    return phase_name

def get_project_name_from(directory_name):
    """Get the project name from the directory name."""

    project_name = ''
    try:
        project_name = substring_from(directory_name, SEPARATOR, 1)
    except IndexError:
        pass

    return project_name

def get_object_name_from(dir_name):
    """Get the object name from the directory name."""

    object_name = ''
    try:
        object_name = substring_from(dir_name, SEPARATOR, 2)
    except IndexError:
        pass

    return object_name

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
    """Write an Excel file from a array of IoT  Dictionnary objects"""

    #Create the Excel File and the Sheets inside.
    book = xlwt.Workbook(encoding="utf-8")
    excel_sheet_main = book.add_sheet("Suivi Global")
    excel_sheet_suivi_evaluation = book.add_sheet("Evaluation")
    excel_sheet_suivi_qualification = book.add_sheet("Qualification")
    excel_sheet_suivi_delivery = book.add_sheet("Delivery Request")

    #Write Headers for every sheets in Excel file.
    write_header_to_sheet(excel_sheet_main)
    write_header_to_sheet(excel_sheet_suivi_evaluation)
    write_header_to_sheet(excel_sheet_suivi_qualification)
    write_header_to_sheet(excel_sheet_suivi_delivery)

    #Rows for each sheet starts at 1 as row 0 is for the header.
    main_sheet_row = 1
    evaluation_sheet_row = 1
    qualification_sheet_row = 1
    delivery_sheet_row = 1
    #Column starts as 0, as the table starts from the left.
    column = 0

    #Write Content
    for object_dictionnary in objects_dictionnary_array:
        write_object_status_to_report(object_dictionnary, excel_sheet_main, main_sheet_row, column)

        if object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['1']:
        #Write the Object to the "Evaluation des Risques" sheet.
            write_object_status_to_report(object_dictionnary,
                                          excel_sheet_suivi_evaluation,
                                          evaluation_sheet_row,
                                          column)
            evaluation_sheet_row = evaluation_sheet_row + 1

        elif object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['2']:
        #Write the Object to the "HLRA - Risk Assessments" sheet.
            write_object_status_to_report(object_dictionnary,
                                          excel_sheet_suivi_qualification,
                                          qualification_sheet_row,
                                          column)
            qualification_sheet_row = qualification_sheet_row + 1

        elif object_dictionnary['SecurityProcessPhase'] == PHASES_DICT['3']:
        #Write the Object to the "Audits de Securite" sheet.
            write_object_status_to_report(object_dictionnary,
                                          excel_sheet_suivi_delivery,
                                          delivery_sheet_row, column)
            delivery_sheet_row = delivery_sheet_row + 1
        main_sheet_row = main_sheet_row + 1

    book.save(file_name)

def populate_projects_array():
    """Populate an array with Dictionnaries objects describing
    the IoT objects from the directory tree"""
    objects_dictionnary_array = []

    #for dirpath, dirnames in os.walk(CURRENT_DIR):
    for dirpath, dirnames, files in os.walk(CURRENT_DIR):

        #Add all "Not Started" Objects to the result Array
        for dirname in dirnames:
            if dirname.lower().startswith(NOT_STARTED_PHASE.lower()):
                current_phase = phase_name_from(dirpath)

                object_iot_dictionnary = {'SecurityProcessPhase':current_phase,
                                          'ProjectName':get_project_name_from(dirname),
                                          'ObjectName': get_object_name_from(dirname),
                                          'ObjectState':NOT_STARTED_PHASE,
                                          'PersonInCharge':'',
                                          'Result':''}
                objects_dictionnary_array.append(object_iot_dictionnary)

        #Add all "In Progress" Objects to the result Array
        for dirname in dirnames:
            if dirname.lower().startswith(IN_PROGRESS_PHASE.lower()):
                current_phase = phase_name_from(dirpath)

                object_iot_dictionnary = {'SecurityProcessPhase':current_phase,
                                          'ProjectName':get_project_name_from(dirname),
                                          'ObjectName': get_object_name_from(dirname),
                                          'ObjectState':IN_PROGRESS_PHASE,
                                          'PersonInCharge':'',
                                          'Result':''}
                objects_dictionnary_array.append(object_iot_dictionnary)

        #Add all "Done" Objects to the result Array
        for dirname in dirnames:
            if dirname.lower().startswith(DONE_PHASE.lower()):
                current_phase = phase_name_from(dirpath)

                object_iot_dictionnary = {'SecurityProcessPhase':current_phase,
                                          'ProjectName':get_project_name_from(dirname),
                                          'ObjectName': get_object_name_from(dirname),
                                          'ObjectState':DONE_PHASE,
                                          'PersonInCharge':'',
                                          'Result':''}
                objects_dictionnary_array.append(object_iot_dictionnary)

    return objects_dictionnary_array
