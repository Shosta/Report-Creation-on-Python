'''
Counter Helpers functions
'''
import variables


def count_b2b_objects(iot_objects_array):
    '''Return the number of b2b object in the iot_object_array'''
    are_b2b_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.is_b2b:
            are_b2b_counter = are_b2b_counter + 1

    return are_b2b_counter


def count_b2c_objects(iot_objects_array):
    '''Return the number of b2c object in the iot_object_array'''
    are_b2c_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.is_b2c:
            are_b2c_counter = are_b2c_counter + 1

    return are_b2c_counter


def count_high_risks(iot_objects_array):
    '''Calculated from the result of the iot_objects that could be 'High Risks Identified'.'''
    are_high_risks_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.result.lower().startswith('high risk'):
            are_high_risks_counter = are_high_risks_counter + 1

    return are_high_risks_counter


def count_light_risks(iot_objects_array):
    '''Calculated from the result of the iot_objects that could be 'Light Risks Identified'.'''
    are_light_risks_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.result.lower().startswith('light risk'):
            are_light_risks_counter = are_light_risks_counter + 1

    return are_light_risks_counter


def count_not_started(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'Not Started' step.'''
    are_not_started_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.security_phase_progress == variables.NOT_STARTED_PHASE:
            are_not_started_counter = are_not_started_counter + 1

    return are_not_started_counter


def count_in_progress(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'In Progress' step.'''
    are_in_progress_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.security_phase_progress == variables.IN_PROGRESS_PHASE:
            are_in_progress_counter = are_in_progress_counter + 1

    return are_in_progress_counter


def count_on_hold(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'On Hold' step.'''
    are_on_hold_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.security_phase_progress == variables.ON_HOLD_PHASE:
            are_on_hold_counter = are_on_hold_counter + 1

    return are_on_hold_counter


def count_done(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'Done' step.'''
    are_done_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.security_phase_progress == variables.DONE_PHASE:
            are_done_counter = are_done_counter + 1

    return are_done_counter


def count_not_evaluated(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'Not Evaluated' step.'''
    are_not_evaluated_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.security_phase_progress == variables.NOT_EVALUATED_PHASE:
            are_not_evaluated_counter = are_not_evaluated_counter + 1

    return are_not_evaluated_counter


def count_objects_in_evaluation(iot_objects_array):
    '''Return the number of ObjectInfo that are in an 'Evaluation' step.'''
    are_in_evaluation_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.delivery_security_process_phase == variables.EVALUATION_PHASE:
            are_in_evaluation_counter = are_in_evaluation_counter + 1

    return are_in_evaluation_counter


def count_objects_in_qualification(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'In Progress' step.'''
    are_in_qualification_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.delivery_security_process_phase == variables.QUALIFICATION_PHASE:
            are_in_qualification_counter = are_in_qualification_counter + 1

    return are_in_qualification_counter


def count_objects_in_audit(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'Security Aduti' step.'''
    are_in_audit_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.delivery_security_process_phase == variables.SECURITY_AUDIT_PHASE:
            are_in_audit_counter = are_in_audit_counter + 1

    return are_in_audit_counter


def count_objects_in_testing(iot_objects_array):
    '''Return the number of ObjectInfo that are in a 'Testing' step.'''
    are_in_testing_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.delivery_security_process_phase == variables.TESTING_PHASE:
            are_in_testing_counter = are_in_testing_counter + 1

    return are_in_testing_counter


def count_objects_on_the_market(iot_objects_array):
    '''Return the number of ObjectInfo that are 'On the market'.'''
    are_on_the_market_counter = 0

    for iot_object in iot_objects_array:
        if iot_object.delivery_security_process_phase == variables.IN_LIFE_PHASE:
            are_on_the_market_counter = are_on_the_market_counter + 1

    return are_on_the_market_counter
