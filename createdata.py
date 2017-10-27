"""
Populate an Array of ObjectInfo that describe IoT objects
 and their current status
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
import variables
from report_creation_utils import xml_utils

PHASES_DICT = {
    '1': '1. Evaluation des Risques',
    '2': '2. HLRA - Risk Assessments',
    '3': '3. Audits de Securite',
    '4': '4. Tests Manuels',
    '5': '5. Legal and Contracts',
    '6': '6. Delivery and Quality'
}


SEPARATOR = ' - '

NOT_STARTED_PROJECTS_COUNTER = 0
IN_PROGRESS_PROJECTS_COUNTER = 0
DONE_PROJECTS_COUNTER = 0


def __substring_from(string, index):
    """Retrieve the string from a larger string splitted regarding the separator
    and returning the indexed one.
    """

    liste = os.path.split(string)
    substring = ""
    try:
        substring = liste[index]
    except IndexError:
        if index-1 >= 0:
            substring = __substring_from(string, index-1)

    return substring


def get_deliveryprocess_name(directory_path):
    """Get the available phases names from the directory tree."""

    phase_name = ''
    try:
        phase_name = __substring_from(directory_path, 1)
    except IndexError:
        pass
    
    if phase_name == PHASES_DICT['1']:
        return variables.EVALUATION_PHASE
    elif phase_name == PHASES_DICT['2']:
        return variables.QUALIFICATION_PHASE
    elif phase_name == PHASES_DICT['3']:
        return variables.SECURITY_AUDIT_PHASE
    elif phase_name == PHASES_DICT['4']:
        return variables.SECURITY_AUDIT_PHASE
    
    return phase_name


def get_project_name(directory_name):
    """Get the project name from the directory name.\n

    As the directory name is commonly name as following :
    'Delivery process state - Project Name - Object Name'
    We split the directory name with the separator ' - ' and get the second
    part to retrieve the project name."""

    project_name = ''
    try:
        project_name = directory_name.split(' - ')[1]
    except IndexError:
        pass

    return project_name


def get_object_name(directory_name):
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


def get_person_in_charge(xml_file):
    '''Return the data in the 'personInCharge' item in the xml file'''

    return xml_utils.get_element_value(xml_file, 'personInCharge')


def get_result(xml_file):
    '''Return the data in the 'result' item in the xml file'''

    return xml_utils.get_element_value(xml_file, 'result')


def get_comment(xml_file):
    '''Return the data in the 'Comments' item in the xml file'''

    return xml_utils.get_element_value(xml_file, 'Comments')


def get_go_no_go_result(xml_file):
    '''Return the data in the 'goNoGo' item in the xml file'''

    return xml_utils.get_element_value(xml_file, 'goNoGo')


def is_b2c_object(xml_file):
    '''Return a boolean that indicates if the object is considered as a B2C one.'''
    return xml_utils.get_attribute_value(xml_file, 'type', 'B2C') == 'true'


def is_others_object(xml_file):
    '''Return a boolean that indicates if the object is considered as a Others one.'''
    return xml_utils.get_attribute_value(xml_file, 'type', 'OTHER') == 'true'


def is_eiot_object(xml_file):
    '''Return a boolean that indicates if the object is considered as a eIoT (enterprise Internet of Things) one.'''
    return xml_utils.get_attribute_value(xml_file, 'type', 'EIOT') == 'true'


def __get_high_risks_counter(xml_file):
    return xml_utils.get_attribute_value(xml_file, 'Risks', 'highRiskCounter')


def __get_moderate_risks_counter(xml_file):
    return xml_utils.get_attribute_value(xml_file, 'Risks', 'moderateRiskCounter')


def __get_light_risks_counter(xml_file):
    return xml_utils.get_attribute_value(xml_file, 'Risks', 'lightRiskCounter')


def __get_accepted_risks_counter(xml_file):
    return xml_utils.get_attribute_value(xml_file, 'Risks', 'acceptedRiskCounter')


def __get_solved_risks_counter(xml_file):
    return xml_utils.get_attribute_value(xml_file, 'Risks', 'solvedRiskCounter')


def create_object_info(directory_path, directory_name, object_state):
    '''
    Return the Object dictionnary from a 'directory_path',
    a 'directory_name'
    and an Object State.

    Params :
    directory_path,
    directory_name,
    object_state,

    Return :
    An ObjectInfo with all its members populated with the appropriate data from the parameters.
    '''
    current_phase = get_deliveryprocess_name(directory_path)
    person_in_charge = ''
    result = ''
    comment = ''
    has_go_from_stakeholders = False
    is_b2b = False
    is_b2c = False
    is_eiot = False

    high_risks_counter = 0
    moderate_risks_counter = 0
    light_risks_counter = 0
    accepted_risks_counter = 0
    solved_risks_counter = 0

    from os import listdir
    from os.path import isfile, join
    onlyfiles = [f for f in listdir(os.path.join(directory_path, directory_name)) if isfile(join(os.path.join(directory_path, directory_name), f))]
    for cur_file in onlyfiles:
        # The information are retrieved from the 'project_info.xml' file.
        if cur_file == 'project_info.xml':
            xml_file = join(directory_path, directory_name, cur_file)
            person_in_charge = get_person_in_charge(xml_file)
            result = get_result(xml_file)
            comment = get_comment(xml_file)

            # Get booleans.
            has_go_from_stakeholders = (get_go_no_go_result(xml_file) == 'Go')
            is_b2b = is_others_object(xml_file)
            is_b2c = is_b2c_object(xml_file)
            is_eiot = is_eiot_object(xml_file)

            # Get Risks counters
            high_risks_counter = __get_high_risks_counter(xml_file)
            moderate_risks_counter = __get_moderate_risks_counter(xml_file)
            light_risks_counter = __get_light_risks_counter(xml_file)
            accepted_risks_counter = __get_accepted_risks_counter(xml_file)
            solved_risks_counter = __get_solved_risks_counter(xml_file)


    from IoTProjectInfo import ObjectInfo
    object_info = ObjectInfo()

    object_info.project_name = get_project_name(directory_name)
    object_info.object_name = get_object_name(directory_name)
    object_info.is_b2c = is_b2c
    object_info.is_b2b = is_b2b
    object_info.is_eiot = is_eiot

    object_info.security_phase_progress = object_state
    object_info.delivery_security_process_phase = current_phase


    object_info.person_in_charge = person_in_charge
    object_info.has_go_from_stakeholders = has_go_from_stakeholders
    object_info.result = result
    object_info.comment = comment

    object_info.high_risks_counter = high_risks_counter
    object_info.moderate_risks_counter = moderate_risks_counter
    object_info.light_risks_counter = light_risks_counter
    object_info.accepted_risks_counter = accepted_risks_counter
    object_info.solved_risks_counter = solved_risks_counter

    return object_info


def populate_objects_array(security_projects_folder):
    '''
    Populate an array with 'ObjectInfo' objects describing the IoT objects from the directory tree.
    '''
    iot_objects_array = []

    # for dirpath, dirnames in os.walk(variables.SECURITY_PROJECTS_DIRPATH):
    for dirpath, dirnames, files in os.walk(security_projects_folder):

        dirnames.sort()
        for dirname in dirnames:
            # Add all "Not Evaluated" Objects to the result Array
            if dirname.lower().startswith(variables.NOT_EVALUATED_PHASE.lower()):
                iot_object = create_object_info(
                    dirpath,
                    dirname,
                    variables.NOT_EVALUATED_PHASE)
                iot_objects_array.append(iot_object)

            # Add all "Not Started" Objects to the result Array
            if dirname.lower().startswith(variables.NOT_STARTED_PHASE.lower()):
                iot_object = create_object_info(
                    dirpath,
                    dirname,
                    variables.NOT_STARTED_PHASE)
                iot_objects_array.append(iot_object)

            # Add all "In Progress" Objects to the result Array
            if dirname.lower().startswith(variables.IN_PROGRESS_PHASE.lower()):
                iot_object = create_object_info(
                    dirpath,
                    dirname,
                    variables.IN_PROGRESS_PHASE)
                iot_objects_array.append(iot_object)

            # Add all "On Hold" Objects to the result Array
            if dirname.lower().startswith(variables.ON_HOLD_PHASE.lower()):
                iot_object = create_object_info(
                    dirpath,
                    dirname,
                    variables.ON_HOLD_PHASE)
                iot_objects_array.append(iot_object)

            # Add all "Done" Objects to the result Array
            if dirname.lower().startswith(variables.DONE_PHASE.lower()):
                iot_object = create_object_info(
                    dirpath,
                    dirname,
                    variables.DONE_PHASE)
                iot_objects_array.append(iot_object)

    return iot_objects_array
