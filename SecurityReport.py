'''
Classes to describe a Security Report object.
'''
import os.path

class Highlights:
    '''
    A Highlighs object.

    It contains a Monthly Highlight, some Highlighted Risks and some Identified Problems.
    '''

    def __init__(self):
        # print('Highlights')
        self._monthly_highlights = ""
        self._highlighted_risks = ""
        self._problems_identified = ""

    @property
    def monthly_highlights(self):
        return self._monthly_highlights

    @property
    def highlighted_risks(self):
        return self._highlighted_risks

    @property
    def problems_identified(self):
        return self._problems_identified

    @monthly_highlights.setter
    def monthly_highlights(self, value):
        self._monthly_highlights = value

    @highlighted_risks.setter
    def highlighted_risks(self, value):
        self._highlighted_risks = value

    @problems_identified.setter
    def problems_identified(self, value):
        self._problems_identified = value

class Risks:
    '''
    A Risks object.

    It contains some Light Risks and some High Risks.
    '''

    def __init__(self):
        # print('Risks')
        self._high_risks = 0
        self._moderate_risks = 0
        self._light_risks = 0
        self._riskiestobjects_list = []

    @property
    def high_risks(self):
        return self._high_risks

    @property
    def moderate_risks(self):
        return self._moderate_risks

    @property
    def light_risks(self):
        return self._light_risks

    @property
    def riskiestobjects_list(self):
        return self._riskiestobjects_list

    @high_risks.setter
    def high_risks(self, value):
        self._high_risks = value

    @moderate_risks.setter
    def moderate_risks(self, value):
        self._moderate_risks = value

    @light_risks.setter
    def light_risks(self, value):
        self._light_risks = value

    @riskiestobjects_list.setter
    def riskiestobjects_list(self, value):
        self._riskiestobjects_list = value

class SecurityPhaseProgress:
    '''
    The progress of a Delivery phase.

    It is described with the following phases :
    1. Not Started
    2. In Progress
    3. Done
    4. Not Evaluated
    '''

    def __init__(self):
        # print('Delivery Phases')
        self._in_progress_objects = 0
        self._done_objects = 0
        self._not_started_objects = 0
        self._not_evaluated_objects = 0

    @property
    def in_progress_objects(self):
        return self._in_progress_objects

    @property
    def done_objects(self):
        return self._done_objects

    @property
    def not_started_objects(self):
        return self._not_started_objects

    @property
    def not_evaluated_objects(self):
        return self._not_evaluated_objects

    @in_progress_objects.setter
    def in_progress_objects(self, value):
        self._in_progress_objects = value

    @done_objects.setter
    def done_objects(self, value):
        self._done_objects = value

    @not_started_objects.setter
    def not_started_objects(self, value):
        self._not_started_objects = value

    @not_evaluated_objects.setter
    def not_evaluated_objects(self, value):
        self._not_evaluated_objects = value

class DeliverySecurityProcess:
    '''
    The phases of the Security delivery process.

    It has basically 5 phases :
    1. Evaluation
    2. Qualification
    3. SecurityAudit
    4. Testing
    5. InLife
    '''
    def __init__(self):
        # print('Security Phases')
        self._evaluation_objects = 0
        self._qualification_objects = 0
        self._security_audit_objects = 0
        self._testing_objects = 0
        self._in_life_objects = 0

    @property
    def evaluation_objects(self):
        return self._evaluation_objects

    @property
    def qualification_objects(self):
        return self._qualification_objects

    @property
    def security_audit_objects(self):
        return self._security_audit_objects

    @property
    def testing_objects(self):
        return self._testing_objects

    @property
    def in_life_objects(self):
        return self._in_life_objects

    @evaluation_objects.setter
    def evaluation_objects(self, value):
        self._evaluation_objects = value

    @qualification_objects.setter
    def qualification_objects(self, value):
        self._qualification_objects = value

    @security_audit_objects.setter
    def security_audit_objects(self, value):
        self._security_audit_objects = value

    @testing_objects.setter
    def testing_objects(self, value):
        self._testing_objects = value

    @in_life_objects.setter
    def in_life_objects(self, value):
        self._in_life_objects = value


class SecurityReport:
    '''
    A complete Security Report.

    It describes, the Risks, and the Highlights.
    It counts the objects according to the Security phase they are in,
    and the Delivery phase they are in.
    '''

    def __init__(self):
        # print('Security Report')
        self._risks = Risks()
        self._highlights = Highlights()
        self._security_phase_progress = SecurityPhaseProgress()
        self._delivery_security_process = DeliverySecurityProcess()

        from datetime import datetime
        now = datetime.now()
        self._report_date = str(now.year) + "-" + str(now.month)
        self._total_objects_count = 0
        self._b2b_objects_count = 0
        self._b2c_objects_count = 0


    @property
    def security_phase_progress(self):
        return self._security_phase_progress

    @property
    def delivery_security_process(self):
        return self._delivery_security_process

    @property
    def risks(self):
        return self._risks

    @property
    def highlights(self):
        return self._highlights

    @property
    def report_date(self):
        return self._report_date

    @property
    def total_objects_count(self):
        return self._total_objects_count

    @property
    def b2b_objects_count(self):
        return self._b2b_objects_count

    @property
    def b2c_objects_count(self):
        return self._b2c_objects_count

    @security_phase_progress.setter
    def security_phase_progress(self, value):
        self._security_phase_progress = value

    @delivery_security_process.setter
    def delivery_security_process(self, value):
        self._delivery_security_process = value

    @risks.setter
    def risks(self, value):
        self._risks = value

    @highlights.setter
    def highlights(self, value):
        self._highlights = value

    @total_objects_count.setter
    def total_objects_count(self, value):
        self._total_objects_count = value

    @b2b_objects_count.setter
    def b2b_objects_count(self, value):
        self._b2b_objects_count = value

    @b2c_objects_count.setter
    def b2c_objects_count(self, value):
        self._b2c_objects_count = value

    def __get_riskiestobjects_list(self, iot_objects_array):
        import variables
        # 1. remove all Evaluation Phase objects.
        for iot_object in iot_objects_array:
            if iot_object.delivery_security_process_phase == variables.EVALUATION_PHASE:
                 iot_objects_array.remove(iot_object)

        #2. Sort the new array by Highest risk counter
        iot_objects_array.sort(key=lambda iot:iot._high_risks_counter, reverse=True)

        #3. Get the five first objects
        if len(iot_objects_array) < 5:
            return iot_objects_array
        else:
            return iot_objects_array[:5]


    def populate_member_values(self, iot_objects_array):
        '''
        Get all the information from the 'iot_objects_array' in one traversal of the array.
        Then we populate the members of the SecurityReport object with the appropriate values.

        Params:
        iot_objects_array The IoT Objects array from which we are going to create the SecurityReport.
        '''

        self.total_objects_count = len(iot_objects_array)

        # Define the counters.
        # Object type : B2B or B2C
        b2b_counter = 0
        b2c_counter = 0
        # Risks
        high_risks_counter = 0
        moderate_risks_counter = 0
        light_risks_counter = 0

        # Security Progress
        not_started_counter = 0
        in_progress_counter = 0
        on_hold_counter = 0
        done_counter = 0
        not_evaluated_counter = 0

        # Delivery Security Phase
        in_evaluation_counter = 0
        in_qualification_counter = 0
        in_audit_counter = 0
        in_testing_counter = 0
        on_the_market_counter = 0

        for iot_object in iot_objects_array:
            # Object type : B2B or B2C
            if iot_object.is_b2b:
                b2b_counter = b2b_counter + 1
            elif iot_object.is_b2c:
                b2c_counter = b2c_counter + 1

            # Risks            
            high_risks_counter = high_risks_counter + iot_object.high_risks_counter
            moderate_risks_counter = moderate_risks_counter + iot_object.moderate_risks_counter
            light_risks_counter = light_risks_counter + iot_object.light_risks_counter

            # Security Progress
            import variables
            if iot_object.security_phase_progress == variables.DONE_PHASE:
                done_counter = done_counter + 1
            elif iot_object.security_phase_progress == variables.IN_PROGRESS_PHASE:
                in_progress_counter = in_progress_counter + 1
            elif iot_object.security_phase_progress == variables.ON_HOLD_PHASE:
                on_hold_counter = on_hold_counter + 1
            elif iot_object.security_phase_progress == variables.NOT_EVALUATED_PHASE:
                not_evaluated_counter = not_evaluated_counter + 1
            elif iot_object.security_phase_progress == variables.NOT_STARTED_PHASE:
                not_started_counter = not_started_counter + 1

            # Delivery Security Phase
            if iot_object.delivery_security_process_phase == variables.EVALUATION_PHASE:
                in_evaluation_counter = in_evaluation_counter + 1
            elif iot_object.delivery_security_process_phase == variables.QUALIFICATION_PHASE:
                in_qualification_counter = in_qualification_counter + 1
            elif iot_object.delivery_security_process_phase == variables.SECURITY_AUDIT_PHASE:
                in_audit_counter = in_audit_counter + 1
            elif iot_object.delivery_security_process_phase == variables.TESTING_PHASE:
                in_testing_counter = in_testing_counter + 1
            elif iot_object.delivery_security_process_phase == variables.IN_LIFE_PHASE:
                on_the_market_counter = on_the_market_counter + 1

        # Object type : B2B or B2C
        self.b2b_objects_count = b2b_counter
        self.b2c_objects_count = b2c_counter

        # Risks : High risks or Light risks
        self.risks.high_risks = high_risks_counter
        self.risks.moderate_risks = moderate_risks_counter
        self.risks.light_risks = light_risks_counter
        self.risks.riskiestobjects_list = self.__get_riskiestobjects_list(iot_objects_array)

        #SecurityPhaseProgress : Not Started, In Progress, On Hold, Done, Not Evaluated
        self.security_phase_progress.not_started_objects = not_started_counter
        self.security_phase_progress.in_progress_objects = in_progress_counter
        self.security_phase_progress.done_objects = done_counter
        self.security_phase_progress.not_evaluated_objects = not_evaluated_counter

        #DeliverySecurityPhase : Evaluation, Qualification, Testing, Audit, In Life
        self.delivery_security_process.evaluation_objects = in_evaluation_counter
        self.delivery_security_process.qualification_objects = in_qualification_counter
        self.delivery_security_process.testing_objects = in_testing_counter
        self.delivery_security_process.security_audit_objects = in_audit_counter
        self.delivery_security_process.in_life_objects = on_the_market_counter

        #Highlights
        from report_creation_utils import xml_utils
        import variables
        highlights_file_path = os.path.join(variables.SECURITY_PROJECTS_DIRPATH, variables.HISTORY_FOLDER_NAME, variables.HIGHLIGHTS_FILE_NAME)
        monthly_highlights = xml_utils.get_element_value(highlights_file_path, 'MonthlyHighlights')
        highlighted_risks = xml_utils.get_element_value(highlights_file_path, 'HighlightedRisks')
        problems_identified = xml_utils.get_element_value(highlights_file_path, 'ProblemsIdentified')
        self.highlights.monthly_highlights = monthly_highlights
        self.highlights.highlighted_risks = highlighted_risks
        self.highlights.problems_identified = problems_identified


    def save_to_file(self):
        # Open the SecurityReport Templates
        from xml.dom import minidom
        import variables
        path = os.path.join(os.path.dirname(__file__),  variables.TEMPLATES_FOLDER_NAME, variables.SECURITY_REPORT_TEMPLATE_FILE_NAME)
        xml_dom = minidom.parse(path)

        from report_creation_utils import xml_utils
        # Replace the SecurityReport node attributes.
        xml_utils.replace_attribute_value(xml_dom, 'SecurityReport', 'date', str(self.report_date))
        xml_utils.replace_attribute_value(xml_dom, 'SecurityReport', 'totalObjects', str(self.total_objects_count))
        xml_utils.replace_attribute_value(xml_dom, 'SecurityReport', 'B2BObjects', str(self.b2b_objects_count))
        xml_utils.replace_attribute_value(xml_dom, 'SecurityReport', 'B2CObjects', str(self.b2c_objects_count))

        # Replace the SecurityPhases(=>DeliverySecurityProcess) childs attributes.
        xml_utils.replace_attribute_value(xml_dom, 'EvaluationObjects', 'counter', str(self.delivery_security_process.evaluation_objects))
        xml_utils.replace_attribute_value(xml_dom, 'QualificationObjects', 'counter', str(self.delivery_security_process.qualification_objects))
        xml_utils.replace_attribute_value(xml_dom, 'SecurityAuditObjects', 'counter', str(self.delivery_security_process.security_audit_objects))
        xml_utils.replace_attribute_value(xml_dom, 'TestingObjects', 'counter', str(self.delivery_security_process.testing_objects))
        xml_utils.replace_attribute_value(xml_dom, 'InLifeObjects', 'counter', str(self.delivery_security_process.in_life_objects))

        # Replace the ObjectsPhases(=>SecurityPhaseProgress) childs attributes.
        xml_utils.replace_attribute_value(xml_dom, 'InProgressObjects', 'counter', str(self.security_phase_progress.in_progress_objects))
        xml_utils.replace_attribute_value(xml_dom, 'DoneObjects', 'counter', str(self.security_phase_progress.done_objects))
        xml_utils.replace_attribute_value(xml_dom, 'NotStartedObjects', 'counter', str(self.security_phase_progress.not_started_objects))
        xml_utils.replace_attribute_value(xml_dom, 'NotEvaluatedObjects', 'counter', str(self.security_phase_progress.not_evaluated_objects))

        # Replace the Risks childs attributes
        xml_utils.replace_attribute_value(xml_dom, 'HighRisks', 'counter', str(self.risks.high_risks))
        xml_utils.replace_attribute_value(xml_dom, 'ModerateRisks', 'counter', str(self.risks.moderate_risks))
        xml_utils.replace_attribute_value(xml_dom, 'LightRisks', 'counter', str(self.risks.light_risks))
        
        # Replace the Highlights childs values
        xml_utils.replace_element_value(xml_dom, 'MonthlyHighlights', str(self.highlights.monthly_highlights))
        xml_utils.replace_element_value(xml_dom, 'HighlightedRisks', str(self.highlights.highlighted_risks))
        xml_utils.replace_element_value(xml_dom, 'ProblemsIdentified', str(self.highlights.problems_identified))
        
        # Encode the str result to bytes in order to write it to a file.
        result = xml_dom.toxml().encode('utf-8')
        # Save the Security Report in a file.
        security_report_path = os.path.join(variables.SECURITY_PROJECTS_DIRPATH,  variables.HISTORY_FOLDER_NAME, str(self.report_date) + variables.SECURITY_REPORT_FILE_NAME)
        file_handle = open(security_report_path, "wb")
        file_handle.write(result)
        file_handle.close()
    