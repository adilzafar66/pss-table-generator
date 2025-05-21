import json
import os
import sys
import pickle
import subprocess
from pathlib import Path
from PyQt5.QtWidgets import QLineEdit
from consts.common import HTTP, HTTPS, ETAP22_PORT, DATAHUB_FILENAME


def save_line_edit_text(file_path: Path, line_edit: QLineEdit, on_error: callable) -> None:
    """
    Saves the text from a QLineEdit widget to a binary file using pickle.

    :param Path file_path: Path where the file will be saved.
    :param QLineEdit line_edit: QLineEdit widget containing the text to save.
    :param callable on_error: Function to call with an Exception object if saving fails.
    """
    try:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        with file_path.open('wb') as f:
            pickle.dump(line_edit.text(), f, protocol=pickle.HIGHEST_PROTOCOL)
    except Exception as ex:
        on_error(ex)


def load_line_edit_text(file_path: Path, line_edit: QLineEdit, on_error: callable) -> None:
    """
    Loads text from a binary file and sets it into a QLineEdit widget.

    :param Path file_path: Path of the file to load.
    :param QLineEdit line_edit: QLineEdit widget where the loaded text will be set.
    :param callable on_error: Function to call with an Exception object if loading fails.
    """
    if not file_path.is_file():
        return
    try:
        with file_path.open('rb') as f:
            text = pickle.load(f)
            if isinstance(text, str):
                line_edit.setText(text)
    except Exception as ex:
        on_error(ex)


def open_file(file_path: Path) -> None:
    """
    Opens a file using the system's default application.

    :param Path file_path: Path to the file to open.
    """
    if not file_path.is_file():
        return

    try:
        if sys.platform == 'win32':
            os.startfile(file_path)
        else:
            opener = 'open' if sys.platform == 'darwin' else 'xdg-open'
            subprocess.call([opener, str(file_path)])
    except Exception as ex:
        print(f"Failed to open file: {ex}")


def split_string_tags(tag_string: str, delimiter: str = ';') -> list[str]:
    """
    Splits a string into a list of tags based on a delimiter, removing empty items and trimming whitespace.

    :param str tag_string: The string containing tags separated by the delimiter.
    :param str delimiter: The character or string used to split tags. Default is ';'.
    :return list[str]: A list of non-empty, trimmed tag strings.
    """
    return [tag.strip() for tag in tag_string.split(delimiter) if tag.strip()]


def get_datahub_info(project_path: str):
    filepath = Path(project_path) / DATAHUB_FILENAME
    try:
        with open(filepath, 'r') as file:
            services = json.load(file)['Services']
    except FileNotFoundError:
        return HTTP, ETAP22_PORT
    port_number = next((s['Port'] for s in services if s['ServiceName'] == 'EtapApi'), None)
    return HTTPS, str(port_number)
