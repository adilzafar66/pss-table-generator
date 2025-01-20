import json
from consts.consts import BASE_REVISION, SC_STUDY_TAG, SC_STUDY_MODE, SC_STUDY_CASE, DEFAULT_SW_CONFIGS
from consts.consts_dd import CONFIG_MAP_INV
from scenario.scenario import Scenario


class ShortCircuitScenario(Scenario):
    """
    The ShortCircuitScenario class is responsible for creating short circuit study scenarios
    within the ETAP environment. It extends the base Scenario class to include specific
    configurations and scenarios related to short circuit studies.
    """

    def __init__(self, url: str):
        """
        Initializes a ShortCircuitScenario instance with the specified ETAP connection port.

        :param str url: local URL for connecting to ETAP datahub.
        """
        super().__init__(url)

    def create_scenarios(self) -> None:
        """
        Creates short circuit study scenarios by iterating over combinations of switching configurations
        Each scenario is uniquely identified by a combination of the study case and switching configuration.
        The generated scenarios are added to the ETAP project and the scenarios XML file is updated.

        - Switching configurations: Derived from the ETAP project configurations.
        """
        switching_configs = json.loads(self.etap.projectdata.getconfigurations())
        switching_configs = list(filter(self.filter_switching_configs, switching_configs))
        for switching_config in switching_configs:
            switching_config_name = CONFIG_MAP_INV.get(switching_config, switching_config)
            scenario_id = SC_STUDY_TAG + '_' + switching_config_name
            self.create_scenario(scenario_id, switching_config, SC_STUDY_MODE, SC_STUDY_CASE, BASE_REVISION,
                                 scenario_id)
            self.scenario_ids.append(scenario_id)
        self.write_scenario_xml()

    @staticmethod
    def filter_switching_configs(switching_config):
        return True if switching_config in DEFAULT_SW_CONFIGS else False
