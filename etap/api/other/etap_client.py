import sys
import re
import json
import subprocess
from pathlib import Path
from .. import application
from .. import projectdata
from .. import scenario
from netifaces import interfaces, ifaddresses, AF_INET


class EtapClient:
    """
    Client connection class used to communicate with ETAP.
    A connection must be established before interacting with ETAP.
    """

    _etap_path = None

    def __init__(self):
        self._is_for_remote_etap = None
        self._ip_address = None
        self.scenario = None
        self.projectdata = None
        self.application = None
        self._token = None
        self._project_name = None
        self._base_address = None

    def connect(self, base_address: str, project_name: str = None) -> None:
        """
        Establishes a connection with ETAP.

        :param str base_address: Base address of the ETAP server
        :param str project_name: Optional project name
        """
        self._base_address = base_address
        self._project_name = project_name
        self._token = self._get_token()
        self.application = application.Application(base_address, self._token, self._project_name)
        self.projectdata = projectdata.ProjectData(base_address, self._token, self._project_name)
        self.scenario = scenario.Scenario(base_address, self._token, self._project_name)
        self._ip_address = self._extract_ip_address_from_base_address(base_address)
        self._is_for_remote_etap = self._is_request_for_remote_machine(self._ip_address)

    def _get_token(self) -> str | None:
        """
        Reads token from keyFile.json based on the project name.

        :return: Token string
        :rtype: str
        """
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            app_path = Path(sys._MEIPASS)
        else:
            app_path = Path(sys.argv[0]).parent

        with open(app_path / 'res' / 'keyFile.json', 'r', encoding='utf-8') as file:
            content_list = json.load(file)

        for entry in content_list:
            if entry.get("projectName") == self._project_name:
                return entry.get("token")

    @staticmethod
    def set_etap_path(path: str) -> None:
        """
        Sets the ETAP executable path.

        :param str path: Path to etaps64.exe
        """
        EtapClient._etap_path = path

    @staticmethod
    def open(project: str = None):
        """
        Opens an ETAP project if the path is set.

        :param str project: Optional project name
        :return: subprocess.Popen instance or False
        """
        if EtapClient._etap_path is None:
            print("Please set the path of etaps64.exe with set_etap_path first")
            return False

        return subprocess.Popen(fr"{EtapClient._etap_path}\etaps64.exe {project}")

    @staticmethod
    def _extract_ip_address_from_base_address(base_address: str) -> str:
        """
        Extracts IP address from the base address.

        :param str base_address: Base address URL
        :return: Extracted IP address
        :rtype: str
        """
        regex = re.compile(f'{re.escape('//')}(.*){re.escape(':')}')
        matches = regex.findall(base_address)
        return matches[0] if matches else ''

    @staticmethod
    def _is_request_for_remote_machine(ip_address: str) -> bool:
        """
        Determines whether this request is for a remote machine running ETAP.

        :param str ip_address: IP address to check
        :return: True if remote, False otherwise
        :rtype: bool
        """
        try:
            for interface in interfaces():
                if len(ifaddresses(interface)) - 1 < AF_INET:
                    continue
                for link in ifaddresses(interface)[AF_INET]:
                    if link['addr'] == ip_address:
                        return False
        except ValueError:
            pass
        return True

    @staticmethod
    def _remove_unwanted_chars(xml_string: str) -> str:
        """
        Removes unwanted Unicode prefix characters.

        :param str xml_string: XML string to clean
        :return: Cleaned XML string
        :rtype: str
        """
        return xml_string[len('ÿþ'):] if xml_string.startswith('ÿþ') else xml_string

    @staticmethod
    def _remove_xml_version_tag(xml_string: str) -> str:
        """
        Removes the XML version tag if present.

        :param str xml_string: XML string
        :return: Modified XML string
        :rtype: str
        """
        return xml_string.replace('<?xml version="1.0" encoding="utf-16"?>', '')

    @staticmethod
    def _decode_unicode_encoded_string(xml_string: str) -> str:
        """
        Decodes a UTF-16 encoded string.

        :param str xml_string: Encoded string
        :return: Decoded string
        :rtype: str
        """
        return xml_string.encode().decode('utf-16')
