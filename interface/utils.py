import os
import sys
import json
import subprocess
from pathlib import Path
from typing import Optional, Any
from consts.common import HTTP, HTTPS, ETAP22_PORT, DATAHUB_FILENAME
from PyQt5.QtWidgets import QLineEdit, QDoubleSpinBox, QCheckBox, QRadioButton, QWidget, QMainWindow


def get_widget_value(widget: QWidget) -> Optional[Any]:
    """
    Retrieves the current value of a given widget.

    :param QWidget widget: The widget to extract the value from.
    :return: The widget's value if supported, otherwise None.
    :rtype: Optional[Any]
    """
    if isinstance(widget, QLineEdit):
        return widget.text()
    elif isinstance(widget, QDoubleSpinBox):
        return widget.value()
    elif isinstance(widget, (QCheckBox, QRadioButton)):
        return widget.isChecked()
    return None


def set_widget_value(widget: QWidget, value: Any) -> Optional[None]:
    """
    Sets the given value to a supported widget.

    :param QWidget widget: The widget to update.
    :param Any value: The value to set.
    :return: None if successful or if widget type is unsupported.
    :rtype: Optional[None]
    """
    if isinstance(widget, QLineEdit):
        return widget.setText(value)
    elif isinstance(widget, QDoubleSpinBox):
        return widget.setValue(value)
    elif isinstance(widget, (QCheckBox, QRadioButton)):
        return widget.setChecked(value)
    return None


def get_all_inputs(main_window: QMainWindow) -> dict:
    """
    Retrieves values from all supported input widgets in the given main window.

    :param QMainWindow main_window: The main window containing input widgets.
    :return: A dictionary mapping object names to their current values.
    :rtype: dict
    """
    values = {}
    for widget in main_window.findChildren(QWidget):
        value = get_widget_value(widget)
        if value is not None and widget.objectName() and 'spinbox' not in widget.objectName():
            values[widget.objectName()] = value
    return values


def set_all_inputs(main_window: QMainWindow, values: dict) -> None:
    """
    Sets values to widgets in the main window based on the provided dictionary.

    :param QMainWindow main_window: The main window containing the widgets.
    :param dict values: A dictionary mapping object names to values to be set.
    """
    for widget in main_window.findChildren(QWidget):
        obj_name = widget.objectName()
        if obj_name in values:
            set_widget_value(widget, values[obj_name])


def save_inputs(main_window: QMainWindow, path: str) -> None:
    """
    Saves the current widget input values to a JSON file.

    :param QMainWindow main_window: The main window containing the input widgets.
    :param str | Path path: The file path to save the JSON data.
    """
    with open(path, 'w') as f:
        json.dump(get_all_inputs(main_window), f)


def load_inputs(main_window: QMainWindow, path: str | Path) -> None:
    """
    Loads widget input values from a JSON file and sets them in the UI.

    :param QMainWindow main_window: The main window to populate with loaded values.
    :param str | Path path: The file path to load the JSON data from.
    """
    with open(path, 'r') as f:
        set_all_inputs(main_window, json.load(f))


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


def get_datahub_info(project_path: str | Path) -> tuple[str, str]:
    """
    Retrieves and returns the datahub settings for a project for API connection.

    :param str | Path project_path: Path to the ETAP project files.
    :return tuple[str, str]: returns the protocol and the port number to connect to datahub.
    """
    filepath = Path(project_path) / DATAHUB_FILENAME
    try:
        with open(filepath, 'r') as file:
            services = json.load(file)['Services']
    except FileNotFoundError:
        return HTTP, ETAP22_PORT
    port_number = next((s['Port'] for s in services if s['ServiceName'] == 'EtapApi'), None)
    return HTTPS, str(port_number)
