import sqlite3
from pathlib import Path
from sqlite3 import Error, OperationalError
from consts.consts_af import indices, ROUND_DIGITS, FT_M_MULTIPLIER


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
        self.ansi_af_data = {}
        self.sql_ansi_af_info = "SELECT Output, Config FROM IAFStudyCase"
        self.sql_ansi_af_bus = (
            "SELECT IDBus, NomlkV, EqType, Orientation, WDistance, FixedBoundary, ResBoundary, "
            "IEnergy, PBoundary, FCT, FCTPD, ArcVaria, ArcI, FCTPDIa, FaultI, FCTPDIf "
            "FROM BusArcFlash WHERE EqType <> 'Other'"
        )
        self.sql_ansi_af_pd = (
            "SELECT ID, NomlkV, Type, Orientation, WDistance, FixedBoundary, ResBoundary, "
            "IEnergy, PBoundary, EnFCT, FCTPD, ArcVaria, EnIa, FCTPDIa, EnIf, FCTPDIf "
            "FROM PDArcFlash WHERE ID <> '' AND Type = 'SPST Switch'"
        )
        self.filepaths = self.get_filepaths(etap_dir, 'AAFS')

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
        cur.execute(self.sql_ansi_af_info)
        af_info = cur.fetchone()

        if af_info:
            af_info = list(af_info)
            af_data = self._fetch_all_data(cur)
            for i in range(len(af_data)):
                af_data[i].insert(indices['rep'] + 1, Path(af_info[0]).stem)
                af_data[i].insert(indices['con'] + 1, af_info[1])
            self._update_ansi_af_data(af_data)

    def _fetch_all_data(self, cur: sqlite3.Cursor) -> list:
        """
        Fetches all arc flash data from the database for both buses and protective devices.

        :param sqlite3.Cursor cur: Cursor object for executing SQL queries on the database.
        :return: A list of arc flash data entries.
        :rtype: list
        """
        af_data = []
        cur.execute(self.sql_ansi_af_bus)
        af_data.extend([list(elem) for elem in cur.fetchall()])
        cur.execute(self.sql_ansi_af_pd)
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
            _energy = entry[indices['ie'] + 1]
            if _id not in self.ansi_af_data or _energy > self.ansi_af_data[_id][indices['ie']]:
                self.ansi_af_data[_id] = _data

    def parse_ansi_af_data(self, use_si_units: bool, exclude_startswith: list[str], exclude_contains: list[list]):
        """
        Parses and processes the extracted ANSI arc flash data, applying filtering and formatting.

        Entries can be excluded based on specific prefixes or contained strings. Numerical values
        are rounded to a predefined precision, and several fields are converted or recalculated.

        :param bool use_si_units: A flag to determine whether to convert some columns to SI units.
        :param list[str] exclude_startswith: List of string prefixes to exclude from the parsed data.
        :param list[str] exclude_contains: List of strings to exclude if contained in entry IDs.
        """
        def filter_func(pair):
            _id, value = pair
            if any(word in _id for word in exclude_contains):
                return False
            if any(_id.startswith(word) for word in exclude_startswith):
                return False
            return True

        conversion_func = self._convert_to_m if use_si_units else self._convert_to_ft
        for key, data in self.ansi_af_data.items():
            data[indices['lab']] = conversion_func(data[indices['lab']])
            data[indices['rab']] = conversion_func(data[indices['rab']])
            data[indices['afb']] = conversion_func(data[indices['afb']])
            data[indices['ecf']] = self._map_electrode_config(data[indices['ecf']])
            data[indices['fct']] = self._convert_cycle_to_sec(data[indices['fct']])
            data[indices['la_var']] = self._calculate_la_var(data[indices['la_var']])

            for i in range(1, len(data)):
                if isinstance(data[i], float):
                    data[i] = round(data[i], ROUND_DIGITS)
            data[0] = round(data[0], ROUND_DIGITS + 1)

        self.ansi_af_data = dict(filter(filter_func, self.ansi_af_data.items()))
        self.ansi_af_data = dict(sorted(self.ansi_af_data.items(),
                                        key=lambda item: item[1][indices['ie']],
                                        reverse=True))

    @staticmethod
    def _calculate_la_var(value: int) -> int:
        """
        Calculates the LaVar (Low Arc Voltage Variation) based on the input value.

        :param int value: Input value for calculation.
        :return: 100 minus the value if it's non-zero, otherwise None.
        :rtype: int
        """
        return None if int(value) == 0 else 100 - value

    @staticmethod
    def _convert_to_ft(value: float) -> str:
        """
        Converts a numeric value in feet and inches to a formatted string.

        :param float value: The value to convert.
        :return: A string in the format "X'Y"", where X is feet and Y is inches.
        :rtype: str
        """
        ft = int(value)
        inch = round((value - ft) * 12)
        if inch == 12:
            ft += 1
            inch = 0
        return f"{ft}'{inch}\""

    @staticmethod
    def _convert_to_m(value: float) -> float:
        """
        Converts a numeric value in feet and inches to meters.

        :param float value: The value to convert.
        :return: value in meters.
        :rtype: float
        """
        return value * FT_M_MULTIPLIER

    @staticmethod
    def _convert_cycle_to_sec(value: float) -> float:
        """
        Converts a time value from cycles to seconds.

        :param float value: The time value in cycles.
        :return: The time value in seconds.
        :rtype: float
        """
        return value / 60

    @staticmethod
    def _map_electrode_config(index: int) -> str:
        """
        Maps an electrode configuration index to its corresponding name.

        :param int index: The index representing an electrode configuration.
        :return: The corresponding electrode configuration name.
        :rtype: str
        """
        electrode_configs = ['VCB', 'VCBB', 'HCB']
        return electrode_configs[index] if 0 <= index < len(electrode_configs) else 'Unknown'

    @staticmethod
    def get_filepaths(input_dir: Path, ext: str) -> list[str]:
        """
        Retrieves all file paths with the specified extension from the given directory.

        :param Path input_dir: Path to the directory to search.
        :param str ext: File extension to filter by.
        :return: List of file paths with the given extension.
        :rtype: list[str]
        """
        filepaths = []
        for path in input_dir.iterdir():
            if path.is_file() and path.suffix == f'.{ext}':
                filepaths.append(str(path))
        return filepaths
