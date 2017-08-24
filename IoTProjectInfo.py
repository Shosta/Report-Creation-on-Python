'''
A bunch of classes that describes a IoT object integrated in a project. 
'''

class ObjectInfo:
    '''
    Some piece of information about an object.

    Describes a IoT object integrated in a project.
    '''

    def __init__(self):
        self._project_name = ""
        self._object_name = ""
        self._delivery_security_process_phase = ""
        self._security_phase_progress = ""
        self._is_b2b = False
        self._is_b2c = False
        self._person_in_charge = ""
        self._has_go_from_stakeholders = False
        self._result = ""
        self._high_risks_counter = 0
        self._moderate_risks_counter = 0
        self._light_risks_counter = 0

    @property
    def project_name(self):
        return self._project_name

    @property
    def object_name(self):
        return self._object_name

    @property
    def delivery_security_process_phase(self):
        '''It can be Evaluation, Qualification, Testing, Audit, In Life.'''
        return self._delivery_security_process_phase

    @property
    def security_phase_progress(self):
        return self._security_phase_progress

    @property
    def is_b2b(self):
        return self._is_b2b

    @property
    def is_b2c(self):
        return self._is_b2c

    @property
    def person_in_charge(self):
        return self._person_in_charge

    @property
    def has_go_from_stakeholders(self):
        '''A boolean value which informs if that object has a Go for the current Security Phase.'''
        return self._has_go_from_stakeholders

    @property
    def result(self):
        return self._result

    @property
    def high_risks_counter(self):
        return self._high_risks_counter

    @property
    def moderate_risks_counter(self):
        return self._moderate_risks_counter

    @property
    def light_risks_counter(self):
        return self._light_risks_counter

    @project_name.setter
    def project_name(self, value):
        self._project_name = value

    @object_name.setter
    def object_name(self, value):
        self._object_name = value

    @delivery_security_process_phase.setter
    def delivery_security_process_phase(self, value):
        '''It can be Evaluation, Qualification, Testing, Audit, In Life.'''
        self._delivery_security_process_phase = value

    @security_phase_progress.setter
    def security_phase_progress(self, value):
        '''It can be Not Started, In Progress, Done or Not Evaluated'''
        self._security_phase_progress = value

    @is_b2b.setter
    def is_b2b(self, boolean):
        self._is_b2b = boolean

    @is_b2c.setter
    def is_b2c(self, boolean):
        self._is_b2c = boolean

    @person_in_charge.setter
    def person_in_charge(self, string):
        self._person_in_charge = string

    @has_go_from_stakeholders.setter
    def has_go_from_stakeholders(self, boolean):
        self._has_go_from_stakeholders = boolean

    @result.setter
    def result(self, string):
        self._result = string

    @high_risks_counter.setter
    def high_risks_counter(self, value):
        self._high_risks_counter = int(value)

    @moderate_risks_counter.setter
    def moderate_risks_counter(self, value):
        self._moderate_risks_counter = int(value)

    @light_risks_counter.setter
    def light_risks_counter(self, value):
        self._light_risks_counter = int(value)
