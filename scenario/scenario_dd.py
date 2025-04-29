import json
from scenario import utils
from scenario.scenario import Scenario
from consts.common import INV_CONFIG_MAP
from consts.tags import DD_STUDY_MODE, DD_STUDY_CASE, DD_STUDY_CASE_IEC, DD_STUDY_MODE_IEC
from consts.tags import DD_TAG, IEC_TAG, BASE_REVISION


class DeviceDutyScenario(Scenario):
    """
    The DeviceDutyScenario class is responsible for creating device duty study scenarios
    within the ETAP environment. It extends the base Scenario class to handle specific
    configurations related to device duty studies.
    """

    def __init__(self, url: str, use_all_sw_configs: bool):
        """
        Initializes a DeviceDutyScenario instance with the specified ETAP connection port.

        :param str url: local URL for connecting to ETAP datahub.
        """
        super().__init__(url)
        self.use_all_sw_configs = use_all_sw_configs

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
        etap_study_cases = utils.get_study_cases(self._etap)

        if DD_STUDY_CASE_IEC in etap_study_cases:
            study_cases.update({DD_STUDY_MODE_IEC: DD_STUDY_CASE_IEC})

        switching_configs = json.loads(self._etap.projectdata.getconfigurations())

        if not self.use_all_sw_configs:
            switching_configs = list(filter(utils.filter_switching_configs, switching_configs))

        for study_mode, study_case in study_cases.items():
            for switching_config in switching_configs:

                # Map the switching configuration to its corresponding name
                switching_config_name = INV_CONFIG_MAP.get(switching_config, switching_config)

                # Get the standard of the current study (ANSI or IEC)
                standard = study_mode.split()[0]

                # Construct the scenario identifier using the study tag and switching configuration name
                scenario_id = DD_TAG + '_' + switching_config_name + (f'_{IEC_TAG}' if standard == IEC_TAG else '')

                # Create the scenario and add its ID to the list
                self.create_scenario(scenario_id, switching_config, study_mode, study_case, BASE_REVISION, scenario_id)
                self.scenario_ids.append(scenario_id)

        # Write the updated scenarios to the XML file
        self.write_scenario_xml()
