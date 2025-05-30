import math
import sqlite3
from parser import utils
from pathlib import Path
from sqlite3 import Error, OperationalError
from consts.filenames import SC_ANSI_EXT
from consts.queries import ANSI_SC_FAULT_QUERY, ANSI_SC_IMP_QUERY
from consts.tags import FAULT_TAG, IMP_TAG, SC_TAG


class ShortCircuitParser:
    def __init__(self, etap_dir: Path):
        self.ansi_sc_data = {}
        self.parsed_ansi_data = {FAULT_TAG: {}, IMP_TAG: {}}
        self.filepaths = utils.get_filepaths(etap_dir, SC_ANSI_EXT, SC_TAG)

    def extract_ansi_data(self):
        """
        Extracts ANSI arc flash data from each SQLite database in the specified directory.

        It attempts to connect to each database, execute the relevant queries, and fetch data
        for further processing. Any errors during database access are logged.
        """
        for filepath in self.filepaths:
            try:
                conn = sqlite3.connect(filepath)
                cur = conn.cursor()
                self._fetch_data(cur, filepath)
                conn.close()
            except (Error, OperationalError) as e:
                print(f"Error with file {filepath}: {e}")

    def _fetch_data(self, cur: sqlite3.Cursor, filepath):
        """
        Fetches study case information and arc flash data from the given database cursor,
        and processes the fetched data to include additional metadata.

        :param sqlite3.Cursor cur: Cursor object for executing SQL queries on the database.
        """
        cur.execute(ANSI_SC_FAULT_QUERY)
        fault_data = cur.fetchall()
        cur.execute(ANSI_SC_IMP_QUERY)
        imp_data = cur.fetchall()
        config = Path(filepath).stem
        self.ansi_sc_data.update({
            config: {
                FAULT_TAG: fault_data,
                IMP_TAG: imp_data
            }
        })

    def parse_ansi_data(self, exclude_startswith: list[str], exclude_contains: list[str], exclude_except: list[str]):
        """
        Parses the extracted ANSI data based on specified criteria, such as excluding specific IDs or
        calculating asymmetrical and symmetrical current values.

        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        """
        for config, modes in self.ansi_sc_data.items():
            config_tags = config.split('_')
            config_id = config_tags[1]
            for mode, entries in modes.items():
                if mode == FAULT_TAG:
                    self.parse_fault_entries(entries, config_id, exclude_startswith, exclude_contains, exclude_except)
                if mode == IMP_TAG:
                    self.parse_imp_entries(entries, config_id, exclude_startswith, exclude_contains, exclude_except)

    def parse_fault_entries(self, entries: list, config: str, exclude_startswith: list[str],
                            exclude_contains: list[str], exclude_except: list[str]):
        """
        Parses ANSI momentary duty data entries based on specified criteria and calculates necessary values.

        :param list entries: List of entries to parse.
        :param str config: Configuration identifier.
        :param list[str] exclude_startswith: List of strings that, if an ID starts with, should exclude the entry.
        :param list[str] exclude_contains: List of substrings that, if present in an ID, should exclude the entry.
        :param list[str] exclude_except: List of substrings that, if present in an ID, should not exclude the entry.
        """
        for entry in entries:
            _id = entry[0]

            if utils.is_exclusion(_id, exclude_startswith, exclude_contains, exclude_except):
                continue

            _voltage = entry[1]
            _real_3ph = entry[2]
            _ima_3ph = entry[3]
            _mag_3ph = entry[4]
            _phase_3ph = math.degrees(math.atan(_ima_3ph / _real_3ph))
            _real_lg = entry[5]
            _ima_lg = entry[6]
            _mag_lg = entry[7]
            _phase_lg = math.degrees(math.atan(_ima_lg / _real_lg))
            _real_ll = entry[8]
            _ima_ll = entry[9]
            _mag_ll = entry[10]
            _phase_ll = math.degrees(math.atan(_ima_ll / _real_ll))
            _real_llg = entry[11]
            _ima_llg = entry[12]
            _mag_llg = entry[13]
            _phase_llg = math.degrees(math.atan(_ima_llg / _real_llg))

            if _id in self.parsed_ansi_data[FAULT_TAG]:
                self.parsed_ansi_data[FAULT_TAG][_id]['Real3Ph'].update({config: _real_3ph})
                self.parsed_ansi_data[FAULT_TAG][_id]['Imag3Ph'].update({config: _ima_3ph})
                self.parsed_ansi_data[FAULT_TAG][_id]['Mag3Ph'].update({config: _mag_3ph})
                self.parsed_ansi_data[FAULT_TAG][_id]['Ph3Ph'].update({config: _phase_3ph})
                self.parsed_ansi_data[FAULT_TAG][_id]['Phasor3Ph'].update({config: [_mag_3ph, _phase_3ph]})
                self.parsed_ansi_data[FAULT_TAG][_id]['RealLG'].update({config: _real_lg})
                self.parsed_ansi_data[FAULT_TAG][_id]['ImagLG'].update({config: _ima_lg})
                self.parsed_ansi_data[FAULT_TAG][_id]['MagLG'].update({config: _mag_lg})
                self.parsed_ansi_data[FAULT_TAG][_id]['PhLG'].update({config: _phase_lg})
                self.parsed_ansi_data[FAULT_TAG][_id]['PhasorLG'].update({config: [_mag_lg, _phase_lg]})
                self.parsed_ansi_data[FAULT_TAG][_id]['RealLL'].update({config: _real_ll})
                self.parsed_ansi_data[FAULT_TAG][_id]['ImagLL'].update({config: _ima_ll})
                self.parsed_ansi_data[FAULT_TAG][_id]['MagLL'].update({config: _mag_ll})
                self.parsed_ansi_data[FAULT_TAG][_id]['PhLL'].update({config: _phase_ll})
                self.parsed_ansi_data[FAULT_TAG][_id]['PhasorLL'].update({config: [_mag_ll, _phase_ll]})
                self.parsed_ansi_data[FAULT_TAG][_id]['RealLLG'].update({config: _real_llg})
                self.parsed_ansi_data[FAULT_TAG][_id]['ImagLLG'].update({config: _ima_llg})
                self.parsed_ansi_data[FAULT_TAG][_id]['MagLLG'].update({config: _mag_llg})
                self.parsed_ansi_data[FAULT_TAG][_id]['PhLLG'].update({config: _phase_llg})
                self.parsed_ansi_data[FAULT_TAG][_id]['PhasorLLG'].update({config: [_mag_llg, _phase_llg]})
                continue

            entry_data = {
                'Voltage': _voltage,
                'Real3Ph': {config: _real_3ph},
                'Imag3Ph': {config: _ima_3ph},
                'Mag3Ph': {config: _mag_3ph},
                'Ph3Ph': {config: _phase_3ph},
                'Phasor3Ph': {config: [_mag_3ph, _phase_3ph]},
                'RealLG': {config: _real_lg},
                'ImagLG': {config: _ima_lg},
                'MagLG': {config: _mag_lg},
                'PhLG': {config: _phase_lg},
                'PhasorLG': {config: [_mag_lg, _phase_lg]},
                'RealLL': {config: _real_ll},
                'ImagLL': {config: _ima_ll},
                'MagLL': {config: _mag_ll},
                'PhLL': {config: _phase_ll},
                'PhasorLL': {config: [_mag_ll, _phase_ll]},
                'RealLLG': {config: _real_llg},
                'ImagLLG': {config: _ima_llg},
                'MagLLG': {config: _mag_llg},
                'PhLLG': {config: _phase_llg},
                'PhasorLLG': {config: [_mag_llg, _phase_llg]},
            }
            self.parsed_ansi_data[FAULT_TAG].update({_id: entry_data})

    def parse_imp_entries(self, entries: list, config: str, exclude_startswith: list[str],
                          exclude_contains: list[str], exclude_except: list[str]):
        for entry in entries:
            _id = entry[0]

            if utils.is_exclusion(_id, exclude_startswith, exclude_contains, exclude_except):
                continue

            _voltage = entry[1]
            _r_pos = entry[2]
            _x_pos = entry[3]
            _z_pos = entry[4]
            _r_neg = entry[5]
            _x_neg = entry[6]
            _z_neg = entry[7]
            _r_zero = entry[8]
            _x_zero = entry[9]
            _z_zero = entry[10]

            if _id in self.parsed_ansi_data[IMP_TAG]:
                self.parsed_ansi_data[IMP_TAG][_id]['RPosOhm'].update({config: _r_pos})
                self.parsed_ansi_data[IMP_TAG][_id]['XPosOhm'].update({config: _x_pos})
                self.parsed_ansi_data[IMP_TAG][_id]['ZPosOhm'].update({config: _z_pos})
                self.parsed_ansi_data[IMP_TAG][_id]['RNegOhm'].update({config: _r_neg})
                self.parsed_ansi_data[IMP_TAG][_id]['XNegOhm'].update({config: _x_neg})
                self.parsed_ansi_data[IMP_TAG][_id]['ZNegOhm'].update({config: _z_neg})
                self.parsed_ansi_data[IMP_TAG][_id]['RZeroOhm'].update({config: _r_zero})
                self.parsed_ansi_data[IMP_TAG][_id]['XZeroOhm'].update({config: _x_zero})
                self.parsed_ansi_data[IMP_TAG][_id]['ZZeroOhm'].update({config: _z_zero})
                continue

            entry_data = {
                'Voltage': _voltage,
                'RPosOhm': {config: _r_pos},
                'XPosOhm': {config: _x_pos},
                'ZPosOhm': {config: _z_pos},
                'RNegOhm': {config: _r_neg},
                'XNegOhm': {config: _x_neg},
                'ZNegOhm': {config: _z_neg},
                'RZeroOhm': {config: _r_zero},
                'XZeroOhm': {config: _x_zero},
                'ZZeroOhm': {config: _z_zero}
            }
            self.parsed_ansi_data[IMP_TAG].update({_id: entry_data})
