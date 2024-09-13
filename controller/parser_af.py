import sqlite3
from pathlib import Path
from sqlite3 import Error, OperationalError
from consts.consts_af import indices, ROUND_DIGITS


class ArcFlashParser:
    """
    Class to handle extraction, parsing, and processing of Arc Flash data from SQLite databases.
    """

    def __init__(self, etap_dir):
        """
        Initializes the ArcFlash class with SQL queries and an empty dictionary for ANSI Arc Flash data.
        """
        self.ansi_af_data = {}
        self.sql_ansi_af_info = "SELECT Output, Config FROM IAFStudyCase"
        self.sql_ansi_af_bus = (
            "SELECT IDBus, NomlkV, EqType, MainPDIsolation, WDistance, FixedBoundary, ResBoundary, "
            "IEnergy, PBoundary, FCT, FCTPD, ArcVaria, ArcI, FCTPDIa, FaultI, FCTPDIf "
            "FROM BusArcFlash WHERE EqType <> 'Other'"
        )
        self.sql_ansi_af_pd = (
            "SELECT ID, NomlkV, Type, MainPDIsolation, WDistance, FixedBoundary, ResBoundary, "
            "IEnergy, PBoundary, EnFCT, FCTPD, ArcVaria, EnIa, FCTPDIa, EnIf, FCTPDIf "
            "FROM PDArcFlash WHERE ID <> '' AND Type = 'SPST Switch'"
        )
        self.filepaths = self.get_filepaths(etap_dir, 'AAFS')

    def extract_ansi_af_data(self):
        """
        Extracts ANSI Arc Flash data from given SQLite database file paths.
        """
        for file_path in self.filepaths:
            try:
                conn = sqlite3.connect(file_path)
                cur = conn.cursor()
                self._fetch_and_process_data(cur)
                conn.close()
            except (Error, OperationalError) as e:
                print(f"Error with file {file_path}: {e}")

    def _fetch_and_process_data(self, cur):
        """
        Fetches and processes data from the database cursor.

        Args:
            cur (sqlite3.Cursor): SQLite cursor object.
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

    def _fetch_all_data(self, cur):
        """
        Fetches all relevant Arc Flash data from the database cursor.

        Args:
            cur (sqlite3.Cursor): SQLite cursor object.

        Returns:
            list: List of fetched data.
        """
        af_data = []
        cur.execute(self.sql_ansi_af_bus)
        af_data.extend([list(elem) for elem in cur.fetchall()])
        cur.execute(self.sql_ansi_af_pd)
        af_data.extend([list(elem) for elem in cur.fetchall()])
        return af_data

    def _update_ansi_af_data(self, af_data):
        """
        Updates the ANSI Arc Flash data dictionary with new data.

        Args:
            af_data (list): List of Arc Flash data entries.
        """
        for entry in af_data:
            _id = entry[0]
            _data = entry[1:]
            _energy = entry[indices['ie'] + 1]
            if _id not in self.ansi_af_data or _energy > self.ansi_af_data[_id][indices['ie']]:
                self.ansi_af_data[_id] = _data

    def parse_ansi_af_data(self, exclude_startswith, exclude_contains):
        """
        Parses the ANSI Arc Flash data to convert units and perform calculations.
        """

        def filter_func(pair):
            _id, value = pair
            if any(word in _id for word in exclude_contains):
                return False
            if any(_id.startswith(word) for word in exclude_startswith):
                return False
            return True

        for key, data in self.ansi_af_data.items():
            data[indices['pdi']] = self._map_electrode_config(data[indices['pdi']])
            data[indices['lab']] = self._convert_to_ft(data[indices['lab']])
            data[indices['rab']] = self._convert_to_ft(data[indices['rab']])
            data[indices['afb']] = self._convert_to_ft(data[indices['afb']])
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
    def _calculate_la_var(value):
        """
        Calculates the la_var value.

        Args:
            value (int): The input value for calculation.

        Returns:
            int: The calculated la_var value or None if input is 0.
        """
        return None if int(value) == 0 else 100 - value

    @staticmethod
    def _convert_to_ft(value):
        """
        Converts a value in feet to a formatted string in feet and inches.

        Args:
            value (float): Value in feet.

        Returns:
            str: Formatted string in feet and inches.
        """
        ft = int(value)
        inch = round((value - ft) * 12)
        return f"{ft}'{inch}\""

    @staticmethod
    def _convert_cycle_to_sec(value):
        """
        Converts cycles to seconds.

        Args:
            value (float): Value in cycles.

        Returns:
            float: Value in seconds.
        """
        return value / 60

    @staticmethod
    def _map_electrode_config(index):
        """
        Maps an index to an electrode configuration string.

        Args:
            index (int): Index of the electrode configuration.

        Returns:
            str: Electrode configuration string.
        """
        electrode_configs = ['None', 'VCB', 'VCBB']
        return electrode_configs[index] if 0 <= index < len(electrode_configs) else 'Unknown'

    @staticmethod
    def get_filepaths(input_dir, ext):
        filepaths = []
        for path in input_dir.iterdir():
            if path.is_file() and path.suffix == f'.{ext}':
                filepaths.append(str(path))
        return filepaths
