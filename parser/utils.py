from pathlib import Path
from sqlite3 import Error, OperationalError, Connection, connect
from consts.multipliers import FT_M_MULTIPLIER
from consts.tags import AF_VCB_CONFIG, AF_VCBB_CONFIG, AF_HCB_CONFIG


def connect_to_sql_file(filepath):
    try:
        conn = connect(filepath)
    except Error:
        raise Error
    return conn.cursor()


def fetch_sql_data(conn: Connection, query):
    cur = conn.cursor()
    try:
        cur.execute(query)
        return cur.fetchall()
    except OperationalError:
        return []


def calculate_la_var(value: int) -> int:
    """
    Calculates the LaVar (Low Arc Voltage Variation) based on the input value.

    :param int value: Input value for calculation.
    :return: 100 minus the value if it's non-zero, otherwise None.
    :rtype: int
    """
    return None if int(value) == 0 else 100 - value


def convert_to_ft(value: float) -> str:
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


def convert_to_m(value: float) -> float:
    """
    Converts a numeric value in feet and inches to meters.

    :param float value: The value to convert.
    :return: value in meters.
    :rtype: float
    """
    return value * FT_M_MULTIPLIER


def convert_cycle_to_sec(value: float) -> float:
    """
    Converts a time value from cycles to seconds.

    :param float value: The time value in cycles.
    :return: The time value in seconds.
    :rtype: float
    """
    return value / 60


def map_electrode_config(index: int) -> str:
    """
    Maps an electrode configuration index to its corresponding name.

    :param int index: The index representing an electrode configuration.
    :return: The corresponding electrode configuration name.
    :rtype: str
    """
    electrode_configs = [AF_VCB_CONFIG, AF_VCBB_CONFIG, AF_HCB_CONFIG]
    return electrode_configs[index] if 0 <= index < len(electrode_configs) else 'Unknown'


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
