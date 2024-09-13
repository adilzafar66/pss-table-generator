import json
from consts.consts import BASE_REVISION
from scenario.scenario import Scenario
from consts.consts_af import AF_STUDY_TAG, CONFIG_MAP, AF_STUDY_MODE, VCB_CONFIG, VCBB_CONFIG


class ArcFlashScenario(Scenario):
    """
    The ArcFlashScenario class is responsible for creating arc flash study scenarios
    within the ETAP environment. It extends the base Scenario class to include specific
    configurations and scenarios related to arc flash studies.
    """

    def __init__(self, port: int = 65358):
        """
        Initializes an ArcFlashScenario instance with the specified ETAP connection port.

        :param int port: The port number for connecting to the ETAP Datahub (default is 65358).
        """
        super().__init__(AF_STUDY_MODE, port)

    def create_scenarios(self) -> None:
        """
        Creates arc flash study scenarios by iterating over combinations of electrode configurations,
        switching configurations, and revision configurations. Each scenario is uniquely identified
        by a combination of the study case, switching configuration, and revision configuration.

        The generated scenarios are added to the ETAP project and the scenarios XML file is updated.

        - Electrode configurations: VCB (Vacuum Circuit Breaker) and VCBB (Vacuum Circuit Breaker Box).
        - Switching configurations: Derived from the ETAP project configurations, excluding 'Ultimate'.
        - Revision configurations: Derived from the ETAP project revisions.
        """
        # Retrieve switching configurations and remove the 'Ultimate' configuration
        switching_configs = json.loads(self.etap.projectdata.getconfigurations())
        rev_configs = json.loads(self.etap.projectdata.getrevisions())
        switching_configs.remove('Ultimate')

        # Define electrode configurations for arc flash studies
        electrode_configs = [VCB_CONFIG, VCBB_CONFIG]

        # Iterate over electrode, switching, and revision configurations to create scenarios
        for electrode_config in electrode_configs:
            for switching_config in switching_configs:
                for rev_config in rev_configs:
                    # Construct the study case name and scenario identifier
                    study_case = AF_STUDY_TAG + '_' + electrode_config
                    rev_config_name = ('_' + rev_config) if rev_config != BASE_REVISION else ''
                    switching_config_name = CONFIG_MAP.get(switching_config, switching_config)
                    scenario_id = study_case + '_' + switching_config_name + rev_config_name

                    # Create the scenario and add its ID to the list
                    self.create_scenario(scenario_id, switching_config, study_case, rev_config, scenario_id)
                    self.scenario_ids.append(scenario_id)

        # Write the updated scenarios to the XML file
        self.write_scenario_xml()
