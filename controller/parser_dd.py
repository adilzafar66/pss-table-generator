import sqlite3
import etap.api
from pathlib import Path
import xml.etree.ElementTree as ET
from sqlite3 import Error, OperationalError
from consts.consts_dd import MV_SWITCHGEAR_MULTIPLIER, LV_SWITCHGEAR_MULTIPLIER
from consts.consts_dd import MODES, MULTIPLIER, ANSI_EXT, IEC_EXT, ANSI_SP_EXT, IEC_SP_EXT


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
        self.ansi_data = {}
        self.iec_data = {}
        self.mode_mom = MODES['Momentary']
        self.mode_int = MODES['Interrupt']
        self.parsed_ansi_data = {self.mode_mom: {}, self.mode_int: {}}
        self.parsed_iec_data = {self.mode_int: {}}

        self.sql_ansi_int = (r"SELECT PDID, kVnom, FaultedBus, PDType, AdjSym, CapAdjInt "
                             r"FROM SCDSumInt WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
        self.sql_ansi_mom = (r"SELECT PDID, kVnom, PDType, kASymm, kAASymm, CapSym, CapAsym "
                             r"FROM SCDSumMom WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
        self.sql_ansi_int_sp = (r"SELECT PDID, kVnom, FaultedBus, PDType, AdjSym, CapAdjInt "
                                r"FROM SCDSumInt1Ph WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
        self.sql_ansi_mom_sp = (r"SELECT PDID, kVnom, PDType, kASymm, kAASymm, CapSym, CapAsym "
                                r"FROM SCDSumMom1Ph WHERE TRIM(PDID) <> '' ORDER BY PDID ASC")
        self.sql_iec_int = (r"SELECT DeviceID, kVnom, FaultedBus, DeviceType, Ibsymm, Ibasymm, DeviceIbsymm, "
                            r"DeviceIbasym FROM SCIEC3phSum WHERE TRIM(DeviceID) <> '' ORDER BY DeviceID ASC")
        self.sql_iec_int_sp = (r"SELECT DeviceID, kVnom, FaultedBus, DeviceType, Ibsymm, Ibasymm, DeviceIbsymm, "
                               r"DeviceIbasym FROM SCIEC1phSum WHERE TRIM(DeviceID) <> '' ORDER BY DeviceID ASC")

        self.ansi_filepaths = self.get_filepaths(etap_dir, ANSI_EXT)
        self.ansi_sp_filepaths = self.get_filepaths(etap_dir, ANSI_SP_EXT)
        self.iec_filepaths = self.get_filepaths(etap_dir, IEC_EXT)
        self.iec_sp_filepaths = self.get_filepaths(etap_dir, IEC_SP_EXT)

    def connect_to_etap(self, port: int):
        """
        Connects to an ETAP project via the given port and retrieves the project data as XML.

        :param int port: The port number to connect to the ETAP API.
        """
        e = etap.api.connect(f'http://localhost:{port}')
        self.tree = ET.fromstring(e.projectdata.getxml())

    def extract_ansi_data(self):
        """
        Extracts ANSI data from SQLite databases and populates the `ansi_data` attribute.
        It processes both momentary and interrupting duties for three-phase and single-phase systems.
        """
        for filepath in self.ansi_filepaths:
            try:
                conn = sqlite3.connect(filepath)
            except Error:
                continue
            cur = conn.cursor()
            try:
                cur.execute(self.sql_ansi_int)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            try:
                cur.execute(self.sql_ansi_mom)
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
                cur.execute(self.sql_ansi_int_sp)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            try:
                cur.execute(self.sql_ansi_mom_sp)
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
                cur.execute(self.sql_iec_int)
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
                cur.execute(self.sql_iec_int)
                int_data = cur.fetchall()
            except OperationalError:
                int_data = []
            config = Path(filepath).stem
            if config not in self.iec_data:
                self.iec_data[config] = {}
                self.iec_data[config][self.mode_int] = []
            self.iec_data[config][self.mode_int] += int_data

    def parse_ansi_data(self, exclude_startswith: list[str], exclude_contains: list[str],
                        calc_sw_asym: bool, calc_swgr_sym: bool):
        """
        Parses the extracted ANSI data based on specified criteria, such as excluding specific IDs or
        calculating asymmetrical and symmetrical current values.

        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param bool calc_sw_asym: Flag to calculate switch asymmetrical values if they are not present.
        :param bool calc_swgr_sym: Flag to calculate switchgear symmetrical values if required.
        """
        for config, modes in self.ansi_data.items():
            config_tags = config.split('_')
            config_id = config_tags[1]
            for mode, entries in modes.items():
                if mode == self.mode_mom:
                    self.parse_ansi_mom_entries(entries, config_id, exclude_startswith, exclude_contains,
                                                calc_sw_asym, calc_swgr_sym)
                if mode == self.mode_int:
                    self.parse_ansi_int_entries(entries, config_id, exclude_startswith, exclude_contains)

    def parse_iec_data(self, exclude_startswith: list[str], exclude_contains: list[str]):
        """
        Parses the extracted IEC data based on specified criteria, such as excluding specific IDs.

        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        """
        for config, modes in self.iec_data.items():
            config_tags = config.split('_')
            config_id = config_tags[1]
            for mode, entries in modes.items():
                if mode == self.mode_int:
                    self.parse_iec_int_entries(entries, config_id, exclude_startswith, exclude_contains)

    def parse_ansi_mom_entries(self, entries: list, config: str, exclude_startswith: list[str],
                               exclude_contains: list[str], calc_sw_asym: bool = True,
                               calc_swgr_sym: bool = True) -> None:
        """
        Parses ANSI momentary duty data entries based on specified criteria and calculates necessary values.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param bool calc_sw_asym: Flag to calculate switch asymmetrical values if they are not present.
        :param bool calc_swgr_sym: Flag to calculate switchgear symmetrical values if required.
        """
        for entry in entries:
            _id = entry[0]
            _voltage = entry[1]
            _type = entry[2]

            if _type.strip() not in ['Panelboard', 'Switchboard',
                                     'SPST Switch', 'SPDT Switch',
                                     'Switchgear']:
                continue

            if any(word in _id for word in exclude_contains):
                continue

            if any(_id.startswith(word) for word in exclude_startswith):
                continue

            if _id in self.parsed_ansi_data[self.mode_mom]:
                self.parsed_ansi_data[self.mode_mom][_id]['Sym'].update({config: entry[3]})
                self.parsed_ansi_data[self.mode_mom][_id]['Asym'].update({config: entry[4]})
                continue

            if _type.endswith('Switch'):
                cap_sym = entry[6]
                cap_asym = entry[5]
                if not cap_asym and calc_sw_asym:
                    cap_asym = cap_sym * MULTIPLIER
            else:
                cap_sym = entry[5]
                cap_asym = entry[6]

            if 'Switchgear' in _type and calc_swgr_sym:
                if _id.startswith('MV') or _voltage > 0.6:
                    cap_sym = cap_asym / MV_SWITCHGEAR_MULTIPLIER
                elif _id.startswith('LV') or _voltage <= 0.6:
                    cap_sym = cap_asym / LV_SWITCHGEAR_MULTIPLIER

            entry_data = {
                'Voltage': _voltage,
                'Type': _type,
                'Sym': {config: entry[3]},
                'Asym': {config: entry[4]},
                'CapSym': cap_sym,
                'CapAsym': cap_asym
            }
            self.parsed_ansi_data[self.mode_mom].update({_id: entry_data})
        self.parsed_ansi_data[self.mode_mom] = dict(sorted(self.parsed_ansi_data[self.mode_mom].items(),
                                                           key=lambda item: item[1]['Type']))

    def parse_ansi_int_entries(self, entries: list, config: str, exclude_startswith: list[str],
                               exclude_contains: list[str]) -> None:
        """
        Parses ANSI interrupting duty data entries based on specified criteria.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        """
        for entry in entries:
            _id = entry[0]
            _voltage = entry[1]
            _bus = entry[2]
            _device = entry[3]

            if any(word in _id for word in exclude_contains):
                continue

            if any(_id.startswith(word) for word in exclude_startswith):
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
                              exclude_contains: list[str]) -> None:
        """
        Parses IEC interrupting duty data entries based on specified criteria.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        """
        for entry in entries:
            _id = entry[0]
            _voltage = entry[1]
            _bus = entry[2]
            _device = entry[3]

            if _device.strip() not in ['CB', 'FUSE']:
                continue

            if any(word in _id for word in exclude_contains):
                continue

            if any(_id.startswith(word) for word in exclude_startswith):
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

    def is_series_rated(self, element_id: str) -> bool:
        """
        Checks if the element with the given ID is series-rated by examining its comment text.

        :param str element_id: The ID of the element to check.
        :return bool: Returns True if the element is series-rated, otherwise False.
        :raises AttributeError: If the element with the given ID is not found.
        """
        if not self.tree:
            return False
        element = self.tree.find(f'.//LAYOUT//*[@ID="{element_id}"]')
        if element is None:
            raise AttributeError(f'No element with ID {element_id} found')
        return True if 'sr' in element.get('CommentText').lower() else False

    def is_assumed(self, element_id: str) -> bool:
        """
        Checks if the element with the given ID is assumed by examining its comment text.

        :param str element_id: The ID of the element to check.
        :return bool: Returns True if the element is assumed, otherwise False.
        :raises AttributeError: If the element with the given ID is not found.
        """
        if not self.tree:
            return False
        element = self.tree.find(f'.//LAYOUT//*[@ID="{element_id}"]')
        if element is None:
            raise AttributeError(f'No element with ID {element_id} found')
        return True if 'assumed' in element.get('CommentText').lower() else False

    def add_series_rated_ratings(self, entries: dict, cap_tag: str, is_mom: bool=False) -> None:
        """
        Adds series-rated ratings to the provided entries if they are series-rated, based on ANSI momentary data.

        :param dict entries: Dictionary of device entries to update with series-rated ratings.
        :param str cap_tag: The key to update in the device entry with the series-rated capacity.
        :param bool is_mom: Flag indicating if the data is momentary (True) or interrupting (False).
        """
        ansi_mom = self.parsed_ansi_data[self.mode_mom]
        for _id, device_data in entries.items():
            is_series_rated = self.is_series_rated(_id)
            device_data['SeriesRated'] = is_series_rated
            if is_series_rated and not is_mom:
                device_bus = device_data['Bus']
                if device_bus in ansi_mom:
                    series_rated_value = ansi_mom[device_bus]['CapSym']
                    device_data[cap_tag] = series_rated_value

    def mark_assumed_equipment(self, entries: dict) -> None:
        """
        Marks equipment as 'Assumed' in the provided entries based on their assumed status.

        :param dict entries: Dictionary of device entries to update with the 'Assumed' status.
        """
        for _id, device_data in entries.items():
            if self.is_assumed(_id):
                device_data['Assumed'] = True

    def process_series_rated_equipment(self) -> None:
        """
        Processes all parsed ANSI and IEC data to identify series-rated equipment and update their ratings.
        """
        ansi_mom = self.parsed_ansi_data[self.mode_mom]
        ansi_int = self.parsed_ansi_data[self.mode_int]
        iec_int = self.parsed_iec_data[self.mode_int]
        self.add_series_rated_ratings(ansi_mom, 'CapSym', is_mom=True)
        self.add_series_rated_ratings(ansi_int, 'CapAdjSym')
        self.add_series_rated_ratings(iec_int, 'CapLbSym')

    def process_assumed_equipment(self) -> None:
        """
        Processes all parsed ANSI and IEC data to mark assumed equipment.
        """
        process_data = [self.parsed_ansi_data[self.mode_mom],
                        self.parsed_ansi_data[self.mode_int],
                        self.parsed_iec_data[self.mode_int]]
        for data in process_data:
            self.mark_assumed_equipment(data)

    @staticmethod
    def get_filepaths(input_dir, ext):
        filepaths = []
        for path in input_dir.iterdir():
            if path.is_file() and path.suffix == f'.{ext}':
                filepaths.append(str(path))
        return filepaths
