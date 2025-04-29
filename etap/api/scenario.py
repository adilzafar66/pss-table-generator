import urllib.parse
import requests
from typing import Optional, Any
from .other import datahub
from .other import api_constants


class Scenario:
    def __init__(self, base_address: str, token: str, project_name: Optional[str] = None):
        """
        Initialize the Scenario instance.

        :param str base_address: The base URL for API requests.
        :param str token: The authentication token.
        :param Optional[str] project_name: The name of the project (optional).
        """
        self._base_address = base_address
        self._token = token
        self._project_name = project_name

    def get_xml_file_path(self) -> str:
        """
        Retrieve the location of the XML file containing scenario information.

        :return str: URL or file path to the XML file.
        """
        try:
            return datahub.get(
                f'{self._base_address}{api_constants.scenario_getxmlfilepath}',
                token=self._token
            )
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get XML file path: {e}") from e

    def get_xml(self) -> str:
        """
        Retrieve XML describing all scenarios for the current project.

        :return str: XML content.
        """
        try:
            return datahub.get(
                f'{self._base_address}{api_constants.scenario_getxml}',
                token=self._token
            )
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get scenario XML: {e}") from e

    def run(
        self,
        scenario_id: str,
        get_online_data: Optional[str] = None,
        what_if_commands: Optional[Any] = None
    ) -> str:
        """
        Run the specified scenario.

        :param str scenario_id: The ID of the scenario to run.
        :param Optional[str] get_online_data: Optional online data toggle.
        :param Optional[Any] what_if_commands: Optional commands for the run.
        :return str: Response from the API.
        """
        try:
            params = {
                'id': scenario_id,
                'getOnlineData': get_online_data or ''
            }
            query_string = urllib.parse.urlencode(params)
            address = f'{self._base_address}{api_constants.scenario_run}?{query_string}'
            return datahub.post(address, what_if_commands, token=self._token)
        except requests.RequestException as e:
            error_message = getattr(e.response, 'text', str(e))
            raise RuntimeError(f"Failed to run scenario: {error_message}") from e

    def create_scenario(
        self,
        scenario_id: str,
        system: str,
        presentation: str,
        revision_name: str,
        config_name: str,
        study_mode: str,
        study_type: str,
        study_case: str,
        output_report: str,
        what_if_commands: Optional[Any] = None
    ) -> bool:
        """
        Create a new scenario using the given parameters.

        :param str scenario_id: ID of the scenario.
        :param str system: System name.
        :param str presentation: Presentation name.
        :param str revision_name: Revision name.
        :param str config_name: Configuration name.
        :param str study_mode: Study mode.
        :param str study_type: Study type.
        :param str study_case: Study case.
        :param str output_report: Output report name.
        :param Optional[Any] what_if_commands: Optional commands for the scenario.
        :return bool: True if scenario is created successfully, False otherwise.
        """
        try:
            param_values = [
                scenario_id, system, presentation, revision_name, config_name,
                study_mode, study_type, study_case, output_report
            ]
            param_keys = [
                'scenarioID', 'system', 'presentation', 'revisionName', 'configName',
                'studyMode', 'studyType', 'studyCase', 'outputReport'
            ]

            encoded_params = {
                key: urllib.parse.quote(str(value), safe='')
                for key, value in zip(param_keys, param_values)
            }

            query_string = urllib.parse.urlencode(encoded_params)
            address = f'{self._base_address}{api_constants.scenario_createscenario}?{query_string}'

            datahub.post(address, what_if_commands, token=self._token)
            return True
        except requests.RequestException as e:
            error_message = getattr(e.response, 'text', str(e))
            print(f"Scenario creation failed: {error_message}")
            return False
