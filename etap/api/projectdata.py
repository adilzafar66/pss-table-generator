import urllib.parse
from typing import Optional, Any
import requests
from .other import datahub
from .other import api_constants


class ProjectData:
    def __init__(self, base_address: str, token: str, project_name: Optional[str] = None):
        """
        Initialize the ProjectData instance.

        :param str base_address: Base URL for API endpoints.
        :param str token: Authorization token.
        :param Optional[str] project_name: Name of the project (optional).
        """
        self._base_address = base_address
        self._token = token
        self._project_name = project_name

    def get_all_element_data(self, element_type: str) -> str:
        """
        Get all data for a specific element type.

        :param str element_type: The element type.
        :return str: Element data in string format.
        """
        try:
            address = (
                f'{self._base_address}{api_constants.projectdata_getallelementdata}?'
                f'elementType={urllib.parse.quote(element_type)}'
            )
            return datahub.get(address, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get all element data: {e}") from e

    def get_configurations(self) -> str:
        """
        Get all configuration names.

        :return str: Configurations.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.projectdata_getconfigurations}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get configurations: {e}") from e

    def get_element_names(self, element_type: str) -> str:
        """
        Get names of elements of a given type.

        :param str element_type: The element type.
        :return str: Comma-separated element names.
        """
        try:
            address = (
                f'{self._base_address}{api_constants.projectdata_getelementnames}?'
                f'elementType={urllib.parse.quote(element_type)}'
            )
            return datahub.get(address, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get element names: {e}") from e

    def get_revisions(self) -> str:
        """
        Get all revision names.

        :return str: Revision names.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.projectdata_getrevisions}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get revisions: {e}") from e

    def get_study_modes_and_cases(self) -> str:
        """
        Get all study modes and study case names.

        :return str: Study modes and cases.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.projectdata_getstudymodesandcases}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get study modes and cases: {e}") from e

    def get_xml(self) -> str:
        """
        Get full XML for the current project.

        :return str: Project XML data.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.projectdata_getxml}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get project XML: {e}") from e

    def get_element_prop(self, element_type: str, element_name: str, field_name: str) -> str:
        """
        Get value of a specific field of a property of an element.

        :param str element_type: Type of the element.
        :param str element_name: Name of the element.
        :param str field_name: Name of the property field.
        :return str: Property value or empty string on failure.
        """
        try:
            encoded = {
                "elementType": urllib.parse.quote(element_type),
                "elementName": urllib.parse.quote(element_name),
                "fieldName": urllib.parse.quote(field_name),
            }
            address = (
                f'{self._base_address}{api_constants.projectdata_getelementprop}?'
                f'{encoded["elementType"]}&elementName={encoded["elementName"]}&fieldName={encoded["fieldName"]}'
            )
            return datahub.get(address, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get element property: {e}") from e

    def get_element_types(self) -> str:
        """
        Get all supported element types.

        :return str: Supported element types.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.projectdata_getelementtypes}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get element types: {e}") from e

    def set_element_prop(self, element_type: str, element_name: str, field_name: str, value: str) -> str:
        """
        Set value of a specific field of an element.

        :param str element_type: Type of the element.
        :param str element_name: Name of the element.
        :param str field_name: Field name to update.
        :param str value: New value to set.
        :return str: API response.
        """
        try:
            encoded = {
                "elementType": urllib.parse.quote(element_type),
                "elementName": urllib.parse.quote(element_name),
                "fieldName": urllib.parse.quote(field_name),
                "value": urllib.parse.quote(value),
            }
            address = (
                f'{self._base_address}{api_constants.projectdata_setelementprop}?'
                f'{encoded["elementType"]}&elementName={encoded["elementName"]}&'
                f'fieldName={encoded["fieldName"]}&value={encoded["value"]}'
            )
            return datahub.post(address, {}, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to set element property: {e}") from e

    def get_study_case_names(self) -> str:
        """
        Get a list of study case names.

        :return str: Study case IDs.
        """
        try:
            return datahub.get(f'{self._base_address}{api_constants.projectdata_getstudycasenames}', token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get study case names: {e}") from e

    def get_study_case(self, study_case_id: str) -> str:
        """
        Get a study case by ID.

        :param str study_case_id: The study case ID.
        :return str: Study case data.
        """
        try:
            address = (
                f'{self._base_address}{api_constants.projectdata_getstudycase}?'
                f'studyCaseId={urllib.parse.quote(study_case_id)}'
            )
            return datahub.get(address, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to get study case: {e}") from e

    def set_study_case(self, xml_element: Any) -> str:
        """
        Add or update a study case.

        :param Any xml_element: XML element data representing the study case.
        :return str: API response.
        """
        try:
            return datahub.post(f'{self._base_address}{api_constants.projectdata_setstudycase}', xml_element, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to set study case: {e}") from e

    def set_elements_props(self, composite_network: str, multiple_element_json: Optional[Any] = None) -> str:
        """
        Set multiple fields of multiple elements at once.

        :param str composite_network: Name of the composite network.
        :param Optional[Any] multiple_element_json: JSON data specifying field updates.
        :return str: API response.
        """
        try:
            encoded_network = urllib.parse.quote(composite_network)
            address = (
                f'{self._base_address}{api_constants.projectdata_setelementsprops}?'
                f'compositeNetwork={encoded_network}'
            )
            return datahub.post(address, multiple_element_json, token=self._token)
        except requests.RequestException as e:
            raise RuntimeError(f"Failed to set element properties: {e}") from e
