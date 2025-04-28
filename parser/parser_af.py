import sqlite3
from parser import utils
from pathlib import Path
from sqlite3 import Error, OperationalError
from consts.common import ROUND_DIGITS
from consts.columns import AF_COL_INDICES
from consts.filenames import AF_ANSI_EXT
from consts.queries import AF_INFO_QUERY, AF_BUS_QUERY, AF_PD_QUERY



class ArcFlashParser:
    """
    The ArcFlashParser class is responsible for extracting, parsing, and processing arc flash data
    from ETAP arc flash study results stored in SQLite databases.
    """

    def __init__(self, etap_dir: Path):
        """
        Initializes the ArcFlashParser instance with the directory containing arc flash study files.

        :param Path etap_dir: Path to the directory containing the arc flash study files (AAFS files).
        """
        self.parsed_ansi_data = {}
        self.filepaths = utils.get_filepaths(etap_dir, AF_ANSI_EXT)

    def extract_ansi_af_data(self):
        """
        Extracts ANSI arc flash data from each SQLite database in the specified directory.

        It attempts to connect to each database, execute the relevant queries, and fetch data
        for further processing. Any errors during database access are logged.
        """
        for file_path in self.filepaths:
            try:
                conn = sqlite3.connect(file_path)
                cur = conn.cursor()
                self._fetch_and_process_data(cur)
                conn.close()
            except (Error, OperationalError) as e:
                print(f"Error with file {file_path}: {e}")

    def _fetch_and_process_data(self, cur: sqlite3.Cursor):
        """
        Fetches study case information and arc flash data from the given database cursor,
        and processes the fetched data to include additional metadata.

        :param sqlite3.Cursor cur: Cursor object for executing SQL queries on the database.
        """
        cur.execute(AF_INFO_QUERY)
        af_info = cur.fetchone()

        if af_info:
            af_info = list(af_info)
            af_data = self._fetch_all_data(cur)
            for i in range(len(af_data)):
                af_data[i].insert(AF_COL_INDICES['rep'] + 1, Path(af_info[0]).stem)
                af_data[i].insert(AF_COL_INDICES['con'] + 1, af_info[1])
            self._update_ansi_af_data(af_data)

    @staticmethod
    def _fetch_all_data(cur: sqlite3.Cursor) -> list:
        """
        Fetches all arc flash data from the database for both buses and protective devices.

        :param sqlite3.Cursor cur: Cursor object for executing SQL queries on the database.
        :return: A list of arc flash data entries.
        :rtype: list
        """
        af_data = []
        cur.execute(AF_BUS_QUERY)
        af_data.extend([list(elem) for elem in cur.fetchall()])
        cur.execute(AF_PD_QUERY)
        af_data.extend([list(elem) for elem in cur.fetchall()])
        return af_data

    def _update_ansi_af_data(self, af_data: list):
        """
        Updates the internal ANSI arc flash data dictionary by retaining the highest incident
        energy value for each entry.

        :param list af_data: List of arc flash data entries to process and update.
        """
        for entry in af_data:
            _id = entry[0]
            _data = entry[1:]
            _energy = entry[AF_COL_INDICES['ie'] + 1]
            if _id not in self.parsed_ansi_data or _energy > self.parsed_ansi_data[_id][AF_COL_INDICES['ie']]:
                self.parsed_ansi_data[_id] = _data

    def parse_ansi_af_data(self, use_si_units: bool, exclude_startswith: list[str],
                           exclude_contains: list[str], exclude_except: list[str]):
        """
        Parses and processes the extracted ANSI arc flash data, applying filtering and formatting.

        Entries can be excluded based on specific prefixes or contained strings. Numerical values
        are rounded to a predefined precision, and several fields are converted or recalculated.

        :param bool use_si_units: A flag to determine whether to convert some columns to SI units.
        :param list[str] exclude_startswith: List of string prefixes to exclude from the parsed data.
        :param list[str] exclude_contains: List of strings to exclude if contained in entry IDs.
        :param list[str] exclude_except: List of strings to not exclude if contained in entry IDs.
        """
        def filter_func(pair):
            _id, value = pair
            if (any(word in _id for word in exclude_contains)
                    and all(word not in _id for word in exclude_except)):
                return False
            if (any(_id.startswith(word) for word in exclude_startswith)
                    and all(not _id.startswith(word) for word in exclude_except)):
                return False
            return True

        conversion_func = utils.convert_to_m if use_si_units else utils.convert_to_ft
        for key, data in self.parsed_ansi_data.items():
            data[AF_COL_INDICES['lab']] = conversion_func(data[AF_COL_INDICES['lab']])
            data[AF_COL_INDICES['rab']] = conversion_func(data[AF_COL_INDICES['rab']])
            data[AF_COL_INDICES['afb']] = conversion_func(data[AF_COL_INDICES['afb']])
            data[AF_COL_INDICES['ecf']] = utils.map_electrode_config(data[AF_COL_INDICES['ecf']])
            data[AF_COL_INDICES['fct']] = utils.convert_cycle_to_sec(data[AF_COL_INDICES['fct']])
            data[AF_COL_INDICES['la_var']] = utils.calculate_la_var(data[AF_COL_INDICES['la_var']])

            for i in range(1, len(data)):
                if isinstance(data[i], float):
                    data[i] = round(data[i], ROUND_DIGITS)
            data[0] = round(data[0], ROUND_DIGITS + 1)

        self.parsed_ansi_data = dict(filter(filter_func, self.parsed_ansi_data.items()))
        self.parsed_ansi_data = dict(sorted(self.parsed_ansi_data.items(),
                                            key=lambda item: item[1][AF_COL_INDICES['ie']],
                                            reverse=True))
