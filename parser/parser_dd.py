import json
import sqlite3
import etap.api
from pathlib import Path
import xml.etree.ElementTree as ET
from sqlite3 import Error, OperationalError
from consts.consts import TYPE_MAP
from consts.consts_dd import MV_SWITCHGEAR_MULTIPLIER, LV_SWITCHGEAR_MULTIPLIER, DD_STUDY_TAG, COMMENT_VAR, BUS
from consts.consts_dd import MODES, ANSI_EXT, IEC_EXT, ANSI_SP_EXT, IEC_SP_EXT
from consts.sql_queries import *


class DeviceDutyParser:
    """
    Parses and processes device duty data from ANSI and IEC standards.
    Handles extraction and parsing of data related to momentary and interrupting duties
    of protection devices in both three-phase and single-phase systems.
    """

    def __init__(self, etap_dir: Path):
        """
        Initializes the DeviceDutyParser with the given ETAP directory.
        Sets up SQL queries, modes, and file paths for ANSI and IEC data.

        :param Path etap_dir: The directory containing ETAP project files.
        """
        self.tree = None
        self._etap = None
        self.comments = {}
        self.ansi_data = {}
        self.iec_data = {}
        self.mode_mom = MODES['Momentary']
        self.mode_int = MODES['Interrupt']
        self.parsed_ansi_data = {self.mode_mom: {}, self.mode_int: {}}
        self.parsed_iec_data = {self.mode_int: {}}
        self.ansi_filepaths = self.get_filepaths(etap_dir, ANSI_EXT)
        self.ansi_sp_filepaths = self.get_filepaths(etap_dir, ANSI_SP_EXT)
        self.iec_filepaths = self.get_filepaths(etap_dir, IEC_EXT)
        self.iec_sp_filepaths = self.get_filepaths(etap_dir, IEC_SP_EXT)

    def connect_to_etap(self, url: str):
        """
        Connects to an ETAP project via the given url and retrieves the project data as XML.

        :param str url: local URL for connecting to the datahub.
        """
        self._etap = etap.api.connect(url)
        version = json.loads(self._etap.application.version())['Version']
        if version.startswith('22'):
            self.tree = ET.fromstring(self._etap.projectdata.getxml())

    def extract_ansi_data(self, use_all_sw_configs: bool):
        """
        Extracts ANSI data from SQLite databases and populates the `ansi_data` attribute.
        It processes both momentary and interrupting duties for three-phase and single-phase systems.

        :param bool use_all_sw_configs: Flag to indicate whether to use all non-default switching configuration files.
        """
        if not use_all_sw_configs:
            self.filter_filepaths()

        for filepath in self.ansi_filepaths:
            try:
                conn = sqlite3.connect(filepath)
            except Error:
                continue
            cur = conn.cursor()
            try:
                cur.execute(ANSI_INT_QUERY)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            try:
                cur.execute(ANSI_MOM_QUERY)
                mom_data = cur.fetchall()
            except OperationalError:
                mom_data = []
            config = Path(filepath).stem
            self.ansi_data.update({
                config: {
                    self.mode_mom: mom_data,
                    self.mode_int: int_data
                }
            })
        for filepath in self.ansi_sp_filepaths:
            try:
                conn = sqlite3.connect(filepath)
            except Error:
                continue
            cur = conn.cursor()
            try:
                cur.execute(ANSI_INT_SP_QUERY)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            try:
                cur.execute(ANSI_MOM_SP_QUERY)
                mom_data = cur.fetchall()
            except OperationalError:
                mom_data = []
            config = Path(filepath).stem
            if config not in self.ansi_data:
                self.ansi_data[config] = {}
                self.ansi_data[config][self.mode_mom] = []
                self.ansi_data[config][self.mode_int] = []
            self.ansi_data[config][self.mode_mom] += mom_data
            self.ansi_data[config][self.mode_int] += int_data

    def extract_iec_data(self):
        """
        Extracts IEC data from SQLite databases and populates the `iec_data` attribute.
        It processes interrupting duties for both three-phase and single-phase systems.
        """
        for filepath in self.iec_filepaths:
            try:
                conn = sqlite3.connect(filepath)
            except Error:
                continue
            cur = conn.cursor()
            try:
                cur.execute(IEC_INT_QUERY)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            config = Path(filepath).stem
            self.iec_data.update({
                config: {
                    self.mode_int: int_data
                }
            })
        for filepath in self.iec_sp_filepaths:
            try:
                conn = sqlite3.connect(filepath)
            except Error:
                continue
            cur = conn.cursor()
            try:
                cur.execute(IEC_INT_QUERY)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            config = Path(filepath).stem
            if config not in self.iec_data:
                self.iec_data[config] = {}
                self.iec_data[config][self.mode_int] = []
            self.iec_data[config][self.mode_int] += int_data

    def parse_ansi_data(self, exclude_startswith: list[str], exclude_contains: list[str],
                        exclude_except: list[str], add_switches: bool):
        """
        Parses the extracted ANSI data based on specified criteria, such as excluding specific IDs or
        calculating asymmetrical and symmetrical current values.

        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        :param bool add_switches: Flag to indicate whether to add switches to the parsed data.
        """
        for config, modes in self.ansi_data.items():
            config_tags = config.split('_')
            config_id = config_tags[1]
            for mode, entries in modes.items():
                if mode == self.mode_mom:
                    self.parse_ansi_mom_entries(entries, config_id, exclude_startswith,
                                                exclude_contains, exclude_except, add_switches)
                if mode == self.mode_int:
                    self.parse_ansi_int_entries(entries, config_id, exclude_startswith,
                                                exclude_contains, exclude_except)

    def parse_iec_data(self, exclude_startswith: list[str], exclude_contains: list[str], exclude_except: list[str]):
        """
        Parses the extracted IEC data based on specified criteria, such as excluding specific IDs.

        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        """
        for config, modes in self.iec_data.items():
            config_tags = config.split('_')
            config_id = config_tags[1]
            for mode, entries in modes.items():
                if mode == self.mode_int:
                    self.parse_iec_int_entries(entries, config_id, exclude_startswith, exclude_contains, exclude_except)

    def parse_ansi_mom_entries(self, entries: list, config: str, exclude_startswith: list[str],
                               exclude_contains: list[str], exclude_except: list[str], add_switches: bool = True):
        """
        Parses ANSI momentary duty data entries based on specified criteria and calculates necessary values.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        :param bool add_switches: Flag to indicate whether to add switches to the parsed data.
        """
        valid_types = ['Panelboard', 'Switchboard', 'Switchgear', 'MCC']
        if add_switches:
            valid_types += ['SPST Switch', 'SPDT Switch']

        for entry in entries:
            _id = entry[0]
            _voltage = entry[1]
            _type = entry[2]

            if _type.strip() not in valid_types:
                continue

            if (any(word in _id for word in exclude_contains)
                    and all(word not in _id for word in exclude_except)):
                continue

            if (any(_id.startswith(word) for word in exclude_startswith)
                    and all(not _id.startswith(word) for word in exclude_except)):
                continue

            if _type.endswith('Switch'):
                _symmetric = 0
            else:
                _symmetric = entry[3]
            _asymmetric = entry[4]

            if _id in self.parsed_ansi_data[self.mode_mom]:
                self.parsed_ansi_data[self.mode_mom][_id]['Sym'].update({config: _symmetric})
                self.parsed_ansi_data[self.mode_mom][_id]['Asym'].update({config: _asymmetric})
                continue

            cap_sym = entry[5]
            cap_asym = entry[6]

            if 'Switchgear' in _type:
                if _id.startswith('MV') or _voltage > 0.6:
                    cap_sym = cap_asym / MV_SWITCHGEAR_MULTIPLIER
                elif _id.startswith('LV') or _voltage <= 0.6:
                    cap_sym = cap_asym / LV_SWITCHGEAR_MULTIPLIER

            entry_data = {
                'Voltage': _voltage,
                'Type': _type,
                'Sym': {config: _symmetric},
                'Asym': {config: _asymmetric},
                'CapSym': cap_sym,
                'CapAsym': cap_asym
            }
            self.parsed_ansi_data[self.mode_mom].update({_id: entry_data})
        self.parsed_ansi_data[self.mode_mom] = dict(sorted(self.parsed_ansi_data[self.mode_mom].items(),
                                                           key=lambda item: item[1]['Type']))

    def parse_ansi_int_entries(self, entries: list, config: str, exclude_startswith: list[str],
                               exclude_contains: list[str], exclude_except: list[str]):
        """
        Parses ANSI interrupting duty data entries based on specified criteria.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        """
        for entry in entries:
            _id = entry[0]
            _voltage = entry[1]
            _bus = entry[2]
            _device = entry[3]

            if (any(word in _id for word in exclude_contains)
                    and all(word not in _id for word in exclude_except)):
                continue

            if (any(_id.startswith(word) for word in exclude_startswith)
                    and all(not _id.startswith(word) for word in exclude_except)):
                continue

            if _id in self.parsed_ansi_data[self.mode_int]:
                self.parsed_ansi_data[self.mode_int][_id]['AdjSym'].update({config: entry[4]})
                continue

            entry_data = {
                'Voltage': _voltage,
                'Bus': _bus,
                'Device': _device,
                'AdjSym': {config: entry[4]},
                'CapAdjSym': entry[5]
            }
            self.parsed_ansi_data[self.mode_int].update({_id: entry_data})
        self.parsed_ansi_data[self.mode_int] = dict(sorted(self.parsed_ansi_data[self.mode_int].items(),
                                                           key=lambda item: item[1]['Bus']))

    def parse_iec_int_entries(self, entries: list, config: str, exclude_startswith: list[str],
                              exclude_contains: list[str], exclude_except: list[str]):
        """
        Parses IEC interrupting duty data entries based on specified criteria.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        """
        for entry in entries:
            _id = entry[0]
            _voltage = entry[1]
            _bus = entry[2]
            _device = entry[3]

            if _device.strip() not in ['CB', 'Fuse']:
                continue

            if (any(word in _id for word in exclude_contains)
                    and all(word not in _id for word in exclude_except)):
                continue

            if (any(_id.startswith(word) for word in exclude_startswith)
                    and all(not _id.startswith(word) for word in exclude_except)):
                continue

            if _id in self.parsed_iec_data[self.mode_int]:
                self.parsed_iec_data[self.mode_int][_id]['LbSym'].update({config: entry[4]})
                self.parsed_iec_data[self.mode_int][_id]['LbAsym'].update({config: entry[5]})
                continue

            entry_data = {
                'Voltage': _voltage,
                'Bus': _bus,
                'Device': _device,
                'LbSym': {config: entry[4]},
                'LbAsym': {config: entry[5]},
                'CapLbSym': entry[6],
                'CapLbAsym': entry[7]
            }
            self.parsed_iec_data[self.mode_int].update({_id: entry_data})
        self.parsed_iec_data[self.mode_int] = dict(sorted(self.parsed_iec_data[self.mode_int].items(),
                                                          key=lambda item: item[1]['Bus']))

    def is_series_rated(self, element_id: str, element_type: str) -> bool:
        """
        Checks if the element with the given ID is series-rated by examining its comment text.

        :param str element_id: The ID of the element to check.
        :param str element_type: The type of the element to check for.
        :return bool: Returns True if the element is series-rated, otherwise False.
        """
        element_comment = self.get_element_comment(element_id, element_type)
        return True if element_comment and 'sr' in element_comment else False

    def is_assumed(self, element_id: str, element_type: str) -> bool:
        """
        Checks if the element with the given ID is assumed by examining its comment text.

        :param str element_id: The ID of the element to check.
        :param str element_type: The type of the element to check for.
        :return bool: Returns True if the element is assumed, otherwise False.
        """
        element_comment = self.get_element_comment(element_id, element_type)
        return True if element_comment and 'assumed' in element_comment else False

    def add_series_rated_ratings(self, entries: dict, cap_tag: str, is_mom: bool = False):
        """
        Adds series-rated ratings to the provided entries if they are series-rated, based on ANSI momentary data.

        :param dict entries: Dictionary of device entries to update with series-rated ratings.
        :param str cap_tag: The key to update in the device entry with the series-rated capacity.
        :param bool is_mom: Flag indicating if the data is momentary (True) or interrupting (False).
        """
        ansi_mom = self.parsed_ansi_data[self.mode_mom]
        for _id, device_data in entries.items():
            element_type = device_data.get('Type') or device_data.get('Device')
            is_series_rated = self.is_series_rated(_id, element_type)
            device_data['SeriesRated'] = is_series_rated
            if is_series_rated and not is_mom:
                device_bus = device_data['Bus']
                if device_bus in ansi_mom:
                    series_rated_value = ansi_mom[device_bus]['CapSym']
                    device_data[cap_tag] = series_rated_value

    def mark_assumed_equipment(self, entries: dict):
        """
        Marks equipment as 'Assumed' in the provided entries based on their assumed status.

        :param dict entries: Dictionary of device entries to update with the 'Assumed' status.
        """
        for _id, device_data in entries.items():
            element_type = device_data.get('Type') or device_data.get('Device')
            if self.is_assumed(_id, element_type):
                device_data['Assumed'] = True

    def get_element_comment(self, element_id: str, element_type: str):
        """
        Gets the comment text attribute of an element.

        :param str element_id: The ID of the element to check.
        :param str element_type: The type of the element to check for.
        :return bool: Returns the value of the comment text attribute in lowercase.
        :raises AttributeError: If the element with the given ID is not found.
        """
        if self.comments.get(element_id):
            return self.comments.get(element_id)

        if not self.tree:
            elem_type = TYPE_MAP.get(element_type, BUS)
            get_element_prop = self._etap.projectdata.getelementprop
            value = json.loads(get_element_prop(elem_type, element_id, COMMENT_VAR))
            if value.get('Value') == 'Invalid element name':
                raise AttributeError(f'No element with ID {element_id} found')
            comment = value.get('Value') and value.get('Value').lower()
        else:
            element = self.tree.find(f'.//LAYOUT//*[@ID="{element_id}"]')
            if element is None:
                raise AttributeError(f'No element with ID {element_id} found')
            comment = element.get('CommentText').lower()

        self.comments.update({element_id: comment})
        return comment

    def process_series_rated_equipment(self):
        """
        Processes all parsed ANSI and IEC data to identify series-rated equipment and update their ratings.
        """
        ansi_mom = self.parsed_ansi_data[self.mode_mom]
        ansi_int = self.parsed_ansi_data[self.mode_int]
        iec_int = self.parsed_iec_data[self.mode_int]
        self.add_series_rated_ratings(ansi_mom, 'CapSym', is_mom=True)
        self.add_series_rated_ratings(ansi_int, 'CapAdjSym')
        self.add_series_rated_ratings(iec_int, 'CapLbSym')

    def process_assumed_equipment(self):
        """
        Processes all parsed ANSI and IEC data to mark assumed equipment.
        """
        self.mark_assumed_equipment(self.parsed_ansi_data[self.mode_mom])
        self.mark_assumed_equipment(self.parsed_ansi_data[self.mode_int])
        self.mark_assumed_equipment(self.parsed_iec_data[self.mode_int])

    @staticmethod
    def get_filepaths(input_dir, ext):
        filepaths = []
        for path in input_dir.iterdir():
            if path.is_file() and path.suffix == f'.{ext}':
                filepaths.append(str(path))
        return filepaths

    def filter_filepaths(self):
        def filter_func(path_str):
            filename = Path(path_str).stem
            if filename.split('_')[0] in DD_STUDY_TAG:
                return True
            return False
        self.ansi_filepaths = list(filter(filter_func, self.ansi_filepaths))
