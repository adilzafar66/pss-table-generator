from pathlib import Path
from PyQt5.QtCore import pyqtSignal
from consts.common import SUBHEAD_ROW
from consts.filenames import SC_FILENAME
from consts.keys import KEYS_SC_FAULT, KEYS_SC_IMP
from consts.tags import FAULT_TAG, IMP_TAG
from exporters.exporter_sc import ShortCircuitExporter
from parser.parser_sc import ShortCircuitParser
from scenario.scenario_sc import ShortCircuitScenario
from worker.worker import Worker
from consts.columns import SC_FAULT_CONST_COLS, SC_FAULT_VAR_COLS, SC_IMP_CONST_COLS, SC_IMP_VAR_COLS


class ShortCircuitWorker(Worker):
    """
    ShortCircuitWorker class for handling short circuit analysis, parsing, and exporting.
    Inherits from the Worker class and provides specialized methods for short circuit operations.
    """

    start_device_duty_process = pyqtSignal()

    def __init__(self, url: str, input_dir_path: Path, output_dir_path: Path, create_scenarios: bool,
                 run_scenarios: bool, exclude_startswith: list[str], exclude_contains: list[str],
                 exclude_except: list[str], create_table: bool, use_all_sw_configs: bool, *args, **kwargs):
        """
        Initializes the ShortCircuitWorker with parameters specific to short circuit analysis.

        :param str url: local URL for connecting to ETAP datahub.
        :param Path input_dir_path: Path to the directory containing input data files.
        :param Path output_dir_path: Path to the directory where output files will be saved.
        :param bool create_scenarios: Flag to indicate whether scenarios should be created.
        :param bool run_scenarios: Flag to indicate whether scenarios should be executed.
        :param list exclude_startswith: List of prefixes for elements to exclude from parsing.
        :param list exclude_contains: List of substrings; elements containing these will be excluded.
        :param list exclude_except: List of substrings; elements containing these will not be excluded.
        :param bool create_table: A flag to determine whether to create an PDF reports.
        :param bool use_all_sw_configs: Flag to indicate whether to use all available switching configurations.
        :param args: Additional arguments for Worker initialization.
        :param kwargs: Additional keyword arguments for Worker initialization.
        """
        super().__init__(input_dir_path, output_dir_path, create_scenarios, run_scenarios, exclude_startswith,
                         exclude_contains, exclude_except, create_table, *args, **kwargs)
        self.scenario_class = lambda: ShortCircuitScenario(url, use_all_sw_configs)

    def execute_data_parsing(self) -> None:
        """
        Executes the parsing of device duty data by using the DeviceDutyParser class.
        Parses both ANSI and IEC data from the input directory and processes series ratings and
        assumed equipment if specified.
        """
        sc_parser = ShortCircuitParser(self.input_dir_path)
        sc_parser.extract_ansi_data()
        sc_parser.parse_ansi_data(self.exclude_startswith, self.exclude_contains, self.exclude_except)
        self.parsed_ansi_data = sc_parser.parsed_ansi_data

    def execute_data_export(self) -> Path:
        """
        Executes the export of parsed device duty data to an Excel workbook.
        Creates headers, inserts data, and formats the sheets for ANSI momentary, ANSI interrupting, and IEC interrupting data.
        Saves the workbook to the output directory.

        :return: The path to the saved Excel workbook.
        :rtype: Path
        """
        sc_exporter = ShortCircuitExporter()
        sc_exporter.set_ansi_data(self.parsed_ansi_data)

        # Create headers for ANSI momentary, ANSI interrupting, and IEC interrupting sheets
        sc_exporter.create_headers(0, SC_FAULT_CONST_COLS, SC_FAULT_VAR_COLS, 'Fault Type')
        sc_exporter.create_headers(1, SC_IMP_CONST_COLS, SC_IMP_VAR_COLS, 'Fault Type')

        # Insert data into the sheets
        sc_exporter.insert_data(0, FAULT_TAG, KEYS_SC_FAULT)
        sc_exporter.insert_data(1, IMP_TAG, KEYS_SC_IMP, round_to=3)

        # Format headers for each sheet
        sc_exporter.format_headers(0)
        sc_exporter.format_headers(1)

        # Apply formatting to each sheet
        sc_exporter.format_sheet(0, SUBHEAD_ROW, len(SC_FAULT_CONST_COLS), len(SC_FAULT_VAR_COLS), 16)
        sc_exporter.format_sheet(1, SUBHEAD_ROW, len(SC_IMP_CONST_COLS), len(SC_IMP_VAR_COLS), 16)

        # Save the workbook
        project_number = self.input_dir_path.stem
        filename = f'{project_number}_{SC_FILENAME}'
        wb_path = Path(self.output_dir_path, filename)
        sc_exporter.save_workbook(wb_path)
        return wb_path

    def start_next_process(self):
        """
        Sends the trigger to start the device duty process after the thread finishes.
        """
        self.start_device_duty_process.emit()
