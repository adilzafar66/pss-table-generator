import json
import os
from typing import Optional
import requests
import urllib.request
import urllib.parse
from .other import datahub
from .other import api_constants


class Application:
    def __init__(self, base_address: str, token: str, project_name: Optional[str] = None):
        """
        Initialize the Application instance.

        :param str base_address: Base URL for API endpoints.
        :param str token: Authorization token.
        :param Optional[str] project_name: Project name (optional).
        """
        self._base_address = base_address
        self._token = token
        self._project_name = project_name

    def project_file(self) -> str:
        """
        Get information about the currently open project.

        :return str: Project file information.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.application_projectfile}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to retrieve project file information: {e}") from e

    def login(self, token):
        """Return True if login successful; False otherwise."""
        try:
            http_location = f'{self._base_address}{api_constants.application_login}?token='
            params = token
            address = http_location + urllib.parse.quote(params)
            response = datahub.post(address, dict())
            response_dict = json.loads(response)
            returned_token = response_dict['Value']

            with open(os.path.dirname(os.path.abspath(__file__)) + "\\other\\keyFile.json", 'r') as f:
                content_list = json.load(f)

            content_list.append({'projectName': self._project_name, 'token': returned_token})

            with open(os.path.dirname(os.path.abspath(__file__)) + "\\other\\keyFile.json", 'w') as f:
                json.dump(content_list, f)

            return True

        except:
            return False

    def version(self) -> str:
        """
        Get the ETAP software version.

        :return str: Version string.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.application_version}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to retrieve version information: {e}") from e

    def get_active_scenario(self) -> str:
        """
        Get the active scenario.

        :return str: Active scenario string.
        """
        return datahub.get(f'{self._base_address}{api_constants.application_getactivescenario}', token=self._token)
