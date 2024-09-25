import json
from consts.consts import BASE_REVISION
from scenario.scenario import Scenario
from consts.consts_dd import CONFIG_MAP_INV, DD_STUDY_TAG, DD_STUDY_MODE, DD_STUDY_CASE, DD_STUDY_CASE_IEC, \
    DD_STUDY_MODE_IEC


class DeviceDutyScenario(Scenario):
    """
    The DeviceDutyScenario class is responsible for creating device duty study scenarios
    within the ETAP environment. It extends the base Scenario class to handle specific
    configurations related to device duty studies.
    """

    def __init__(self, port: int = 65358):
        """
        Initializes a DeviceDutyScenario instance with the specified ETAP connection port.

        :param port: The port number for connecting to the ETAP Datahub (default is 65358).
        """
        super().__init__(port)

    def create_scenarios(self) -> None:
        """
        Creates device duty study scenarios by iterating over available switching configurations.
        Each scenario is uniquely identified by the device duty study tag and the switching configuration name.

        The generated scenarios are added to the ETAP project, and the scenarios XML file is updated.

        - Switching configurations: Derived from the ETAP project configurations.
        - Study case: Derived from the ETAP project study cases.
        - Revision: Set to BASE_REVISION.
        """

        study_cases = {DD_STUDY_MODE: DD_STUDY_CASE}
        etap_study_cases = json.loads(self.etap.projectdata.getstudycasenames())['NonDefault']
        if DD_STUDY_CASE_IEC in etap_study_cases:
            study_cases.update({DD_STUDY_MODE_IEC: DD_STUDY_CASE_IEC})
        switching_configs = json.loads(self.etap.projectdata.getconfigurations())

        for study_mode, study_case in study_cases.items():
            # Iterate over each switching configuration to create scenarios
            for switching_config in switching_configs:
                # Map the switching configuration to its corresponding name
                switching_config_name = CONFIG_MAP_INV.get(switching_config, switching_config)

                # Get the standard of the current study (ANSI or IEC)
                standard = study_mode.split()[0]

                # Construct the scenario identifier using the study tag and switching configuration name
                scenario_id = DD_STUDY_TAG + '_' + switching_config_name + ('_IEC' if standard == 'IEC' else '')

                # Create the scenario and add its ID to the list
                self.create_scenario(scenario_id, switching_config, study_mode, study_case, BASE_REVISION, scenario_id)
                self.scenario_ids.append(scenario_id)

        # Write the updated scenarios to the XML file
        self.write_scenario_xml()
