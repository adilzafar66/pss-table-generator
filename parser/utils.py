from pathlib import Path
from consts.multipliers import FT_M_MULTIPLIER
from sqlite3 import Error, OperationalError, Cursor, connect
from consts.tags import AF_VCB_CONFIG, AF_VCBB_CONFIG, AF_HCB_CONFIG


def connect_to_sql_file(filepath: str | Path) -> Cursor:
    """
    Connects to a SQLite database file and returns a cursor object.

    :param str | Path filepath: Path to the SQLite database file.
    :return: Cursor object for executing SQL commands.
    :rtype: Cursor
    :raises Error: If the connection to the database fails.
    """
    try:
        conn = connect(filepath)
    except Error:
        raise Error
    return conn.cursor()


def fetch_sql_data(cur: Cursor, query: str) -> list:
    """
    Executes a SQL query using the provided cursor and returns the fetched results.

    :param Cursor cur: Cursor object connected to the SQLite database.
    :param str query: SQL query to execute.
    :return: List of rows returned by the query; empty list if an error occurs.
    :rtype: list
    """
    try:
        cur.execute(query)
        return cur.fetchall()
    except OperationalError:
        return []


def is_exclusion(_id: str, exclude_startswith: list[str], exclude_contains: list[str],
                 exclude_except: list[str]) -> None | bool:
    """
    Determines whether a given ID should be excluded based on matching rules.

    :param str _id: The string ID to evaluate.
    :param list[str] exclude_startswith: List of prefixes to check for exclusion.
    :param list[str] exclude_contains: List of substrings to check for exclusion.
    :param list[str] exclude_except: List of exceptions to override exclusion.
    :return: True if the ID matches exclusion criteria, False otherwise.
    :rtype: bool
    """
    if (any(word in _id for word in exclude_contains)
            and all(word not in _id for word in exclude_except)):
        return True

    if (any(_id.startswith(word) for word in exclude_startswith)
            and all(not _id.startswith(word) for word in exclude_except)):
        return True


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


def get_filepaths(input_dir: Path, ext: str, study_tag: str) -> list[str]:
    """
    Retrieves all file paths with the specified extension from the given directory.

    :param Path input_dir: Path to the directory to search.
    :param str ext: File extension to filter by.
    :param str study_tag: Tag to filter only the study files.
    :return: List of file paths with the given extension.
    :rtype: list[str]
    """
    filepaths = []
    for path in input_dir.iterdir():
        if path.is_file() and path.suffix == f'.{ext}':
            filepaths.append(str(path))
    return filepaths
# and path.stem.startswith(f'{study_tag}_')