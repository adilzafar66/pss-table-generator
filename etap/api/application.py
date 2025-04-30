from typing import Optional
import requests
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
