from pathlib import Path
from PyQt5.QtCore import pyqtSignal
from consts.common import SUBHEAD_ROW
from worker.worker import Worker
from parser.parser_dd import DeviceDutyParser
from exporters.exporter_dd import DeviceDutyExporter
from scenario.scenario_dd import DeviceDutyScenario
from consts.filenames import DD_FILENAME
from consts.tags import MOM_TAG, INT_TAG
from consts.keys import KEYS_DD_MOM, KEYS_DD_INT, KEYS_DD_INT_IEC
from consts.columns import DD_MOM_CONST_COLS, DD_INT_CONST_COLS, DD_MOM_VAR_COLS, DD_INT_VAR_COLS, DD_INT_IEC_VAR_COLS


class DeviceDutyWorker(Worker):
    """
    DeviceDutyWorker class for handling device duty analysis, parsing, and exporting.
    Inherits from the Worker class and provides specialized methods for device duty operations.
    """

    start_arc_flash_process = pyqtSignal()

    def __init__(self, url: str, input_dir_path: Path, output_dir_path: Path, create_scenarios: bool,
                 run_scenarios: bool, exclude_startswith: list[str], exclude_contains: list[str],
                 exclude_except: list[str], create_table: bool, add_switches: bool, use_all_sw_configs: bool,
                 add_series_ratings: bool, mark_assumed: bool, *args, **kwargs):
        """
        Initializes the DeviceDutyWorker with parameters specific to device duty analysis.

        :param str url: local URL for connecting to ETAP datahub.
        :param Path input_dir_path: Path to the directory containing input data files.
        :param Path output_dir_path: Path to the directory where output files will be saved.
        :param bool create_scenarios: Flag to indicate whether scenarios should be created.
        :param bool run_scenarios: Flag to indicate whether scenarios should be executed.
        :param list exclude_startswith: List of prefixes for elements to exclude from parsing.
        :param list exclude_contains: List of substrings; elements containing these will be excluded.
        :param list exclude_except: List of substrings; elements containing these will not be excluded.
        :param bool create_table: A flag to determine whether to create an Excel table.
        :param bool add_switches: Flag to indicate whether to add switches to the Device Duty report.
        :param bool use_all_sw_configs: Flag to indicate whether to use all available switching configurations.
        :param bool add_series_ratings: Flag to indicate if series ratings should be added.
        :param bool mark_assumed: Flag to indicate if assumed equipment should be marked.
        :param args: Additional arguments for Worker initialization.
        :param kwargs: Additional keyword arguments for Worker initialization.
        """
        super().__init__(input_dir_path, output_dir_path, create_scenarios, run_scenarios, exclude_startswith,
                         exclude_contains, exclude_except, create_table, *args, **kwargs)
        self.datahub_url = url
        self.add_switches = add_switches
        self.add_series_ratings = add_series_ratings
        self.mark_assumed = mark_assumed
        self.scenario_class = lambda: DeviceDutyScenario(url, use_all_sw_configs)

    def execute_data_parsing(self) -> None:
        """
        Executes the parsing of device duty data by using the DeviceDutyParser class.
        Parses both ANSI and IEC data from the input directory and processes series ratings and
        assumed equipment if specified.
        """
        dd_parser = DeviceDutyParser(self.input_dir_path)
        dd_parser.extract_ansi_data()
        dd_parser.parse_ansi_data(self.exclude_startswith, self.exclude_contains,
                                  self.exclude_except, self.add_switches)
        dd_parser.extract_iec_data()
        dd_parser.parse_iec_data(self.exclude_startswith, self.exclude_contains, self.exclude_except)
        if self.add_series_ratings or self.mark_assumed:
            dd_parser.connect_to_etap(self.datahub_url)
            if self.add_series_ratings:
                dd_parser.process_series_rated_equipment()
            if self.mark_assumed:
                dd_parser.process_assumed_equipment()

        self.parsed_ansi_data = dd_parser.parsed_ansi_data
        self.parsed_iec_data = dd_parser.parsed_iec_data

    def execute_data_export(self) -> Path:
        """
        Executes the export of parsed device duty data to an Excel workbook. Creates headers, inserts data,
        and formats the sheets for ANSI momentary, ANSI interrupting, and IEC interrupting data. Saves the workbook
        to the output directory.

        :return: The path to the saved Excel workbook.
        :rtype: Path
        """
        dd_exporter = DeviceDutyExporter()
        dd_exporter.set_ansi_data(self.parsed_ansi_data)
        dd_exporter.set_iec_data(self.parsed_iec_data)

        # Create headers for ANSI momentary, ANSI interrupting, and IEC interrupting sheets
        dd_exporter.create_headers(0, DD_MOM_CONST_COLS, DD_MOM_VAR_COLS, 'Momentary Duty')
        dd_exporter.create_headers(1, DD_INT_CONST_COLS, DD_INT_VAR_COLS, 'Interrupting Duty')
        dd_exporter.create_headers(2, DD_INT_CONST_COLS, DD_INT_IEC_VAR_COLS, 'Interrupting Duty')

        # Insert data into the sheets
        dd_exporter.insert_data(0, MOM_TAG, KEYS_DD_MOM)
        dd_exporter.insert_data(1, INT_TAG, KEYS_DD_INT)
        dd_exporter.insert_data(2, INT_TAG, KEYS_DD_INT_IEC, 'iec')

        # Format headers for each sheet
        dd_exporter.format_headers(0)
        dd_exporter.format_headers(1)
        dd_exporter.format_headers(2)

        # Apply formatting to each sheet
        dd_exporter.format_sheet(0, SUBHEAD_ROW, len(DD_MOM_CONST_COLS), len(DD_MOM_VAR_COLS), 16)
        dd_exporter.format_sheet(1, SUBHEAD_ROW, len(DD_INT_CONST_COLS), len(DD_INT_VAR_COLS), 22)
        dd_exporter.format_sheet(2, SUBHEAD_ROW, len(DD_INT_CONST_COLS), len(DD_INT_IEC_VAR_COLS), 16)

        # Save the workbook
        project_number = self.input_dir_path.stem
        filename = f'{project_number}_{DD_FILENAME}'
        wb_path = Path(self.output_dir_path, filename)
        dd_exporter.save_workbook(wb_path)
        return wb_path

    def start_next_process(self):
        """
        Sends the trigger to start the arc flash process after the thread finishes.
        """
        self.start_arc_flash_process.emit()
