"""
Populate an Array of Dictionnary Objects that describe IoT objects and their current status
in terms of security process.\n

A Dictionnary has the following key :\n
'1. Evaluation des Risques',
'2. HLRA - Risk Assessments',
'3. Audits de Securite',
'4. Tests Manuels',
'5. Legal and Contracts',
'6. Delivery and Quality'
"""

import os.path

PHASES_DICT = {
    '1': '1. Evaluation des Risques',
    '2': '2. HLRA - Risk Assessments',
    '3': '3. Audits de Securite',
    '4': '4. Tests Manuels',
    '5': '5. Legal and Contracts',
    '6': '6. Delivery and Quality'
}

NOT_EVALUATED_PHASE = 'Not Evaluated'
NOT_STARTED_PHASE = 'Not Started'
IN_PROGRESS_PHASE = 'In Progress'
DONE_PHASE = 'Done'
SEPARATOR = ' - '
CURRENT_DIR = '..\\.'

NOT_STARTED_PROJECTS_COUNTER = 0
IN_PROGRESS_PROJECTS_COUNTER = 0
DONE_PROJECTS_COUNTER = 0


def substring_from(string, index):
    """Retrieve the string from a larger string splitted regarding the separator
    and returning the indexed one.
    """

    liste = os.path.split(string)
    substring = ""
    try:
        substring = liste[index]
    except IndexError:
        if index-1 >= 0:
            substring = substring_from(string, index-1)

    return substring


def phase_name_from(directory_path):
    """Get the available phases names from the directory tree."""

    phase_name = ''
    try:
        phase_name = substring_from(directory_path, 1)
    except IndexError:
        pass
    return phase_name


def get_project_name_from(directory_name):
    """Get the project name from the directory name.\n

    As the directory name is commonly name as following :
    'Phase State - Project Name - Object Name'
    We split the directory name with the separator ' - ' and get the second
    part to retrieve the project name."""

    project_name = ''
    try:
        project_name = directory_name.split(' - ')[1]
    except IndexError:
        pass

    return project_name


def get_object_name_from(directory_name):
    """Get the object name from the directory name.\n

    As the directory name is commonly name as following :
    'Phase State - Project Name - Object Name'
    We split the directory name with the separator ' - ' and get the third part
    and get the second part to retrieve the object name."""

    object_name = ''
    try:
        object_name = directory_name.split(' - ')[2]
    except IndexError:
        pass

    return object_name


def get_xml_data(xml_file, element_tag_name):
    '''Return the first data of the 'element_tag_name' type item in the
    'xml_file'.'''

    from xml.dom import minidom
    xmldoc = minidom.parse(xml_file)
    element = xmldoc.getElementsByTagName(element_tag_name)

    first_data = ''
    try:
        first_data = list(element)[0].firstChild.data
    except:
        pass

    return first_data


def get_person_in_charge(xml_file):
    '''Return the data in the 'personInCharge' item in the xml file'''

    return get_xml_data(xml_file, 'personInCharge')


def get_result(xml_file):
    '''Return the data in the 'result' item in the xml file'''

    return get_xml_data(xml_file, 'result')


def get_go_no_go_result(xml_file):
    '''Return the data in the 'goNoGo' item in the xml file'''

    return get_xml_data(xml_file, 'goNoGo')


def get_object_from(directory_path, directory_name, object_state):
    '''Return the Object dictionnary from a 'directory_path',
    a 'directory_name'
    and an Object State.'''
    current_phase = phase_name_from(directory_path)

    person_in_charge = ''
    go_no_go_result = ''
    result = ''

    from os import listdir
    from os.path import isfile, join
    onlyfiles = [f for f in listdir(os.path.join(directory_path, directory_name)) if isfile(join(os.path.join(directory_path, directory_name), f))]
    for cur_file in onlyfiles:
        if cur_file == 'project_info.xml':
            xml_file = join(directory_path, directory_name, cur_file)
            person_in_charge = get_person_in_charge(xml_file)
            result = get_result(xml_file)
            go_no_go_result = get_go_no_go_result(xml_file)

    object_iot_dictionnary = {
        'SecurityProcessPhase': current_phase,
        'ProjectName': get_project_name_from(directory_name),
        'ObjectName': get_object_name_from(directory_name),
        'ObjectState': object_state,
        'PersonInCharge': person_in_charge,
        'Result': result,
        'GoNoGo': go_no_go_result
        }

    return object_iot_dictionnary


def populate_projects_array():
    """Populate an array with Dictionnaries objects describing
    the IoT objects from the directory tree"""
    objects_dictionnary_array = []

    # for dirpath, dirnames in os.walk(CURRENT_DIR):
    for dirpath, dirnames, files in os.walk(CURRENT_DIR):

        for dirname in dirnames:
            # Add all "Not Evaluated" Objects to the result Array
            if dirname.lower().startswith(NOT_EVALUATED_PHASE.lower()):
                object_iot_dictionnary = get_object_from(dirpath, dirname, NOT_EVALUATED_PHASE)
                objects_dictionnary_array.append(object_iot_dictionnary)

            # Add all "Not Started" Objects to the result Array
            if dirname.lower().startswith(NOT_STARTED_PHASE.lower()):
                object_iot_dictionnary = get_object_from(dirpath, dirname, NOT_STARTED_PHASE)
                objects_dictionnary_array.append(object_iot_dictionnary)

            # Add all "In Progress" Objects to the result Array
            if dirname.lower().startswith(IN_PROGRESS_PHASE.lower()):
                object_iot_dictionnary = get_object_from(dirpath, dirname, IN_PROGRESS_PHASE)
                objects_dictionnary_array.append(object_iot_dictionnary)

            # Add all "Done" Objects to the result Array
            if dirname.lower().startswith(DONE_PHASE.lower()):
                object_iot_dictionnary = get_object_from(dirpath, dirname, DONE_PHASE)
                objects_dictionnary_array.append(object_iot_dictionnary)

    return objects_dictionnary_array
