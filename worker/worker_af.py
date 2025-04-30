from pathlib import Path
from consts.columns import AF_CONST_COLS
from consts.common import HEADER_ROW
from consts.filenames import AF_FILENAME
from consts.styles import WIDTH_COL_LRG
from parser.parser_af import ArcFlashParser
from exporters.exporter_af import ArcFlashExporter
from scenario.scenario_af import ArcFlashScenario
from worker.worker import Worker


class ArcFlashWorker(Worker):
    """
    ArcFlashWorker class for handling the creation, parsing, and exporting of arc flash data.
    Inherits from Worker class and overrides specific methods for arc flash data processing.
    """

    def __init__(self, url: str, input_dir_path: Path, output_dir_path: Path, create_scenarios: bool,
                 run_scenarios: bool, exclude_startswith: list[str], exclude_contains: list[str],
                 exclude_except: list[str], create_table: bool, use_si_units: bool, high_energy: float,
                 low_energy: float, revisions: list[str] | None = None, *args, **kwargs):
        """
        Initializes the ArcFlashWorker with specific parameters for arc flash processing.

        :param str url: local URL for connecting to ETAP datahub.
        :param Path input_dir_path: Path to the input directory containing data.
        :param Path output_dir_path: Path to the output directory for saving results.
        :param bool create_scenarios: Flag to indicate if scenarios should be created.
        :param bool run_scenarios: Flag to indicate if scenarios should be run.
        :param list exclude_startswith: List of strings; elements starting with these prefixes will be excluded.
        :param list exclude_contains: List of strings; elements containing these substrings will be excluded.
        :param list exclude_except: List of substrings; elements containing these will not be excluded.
        :param bool create_table: A flag to determine whether to create an Excel table.
        :param bool use_si_units: A flag to determine whether to convert some columns to SI units.
        :param float high_energy: Threshold value for high energy highlighting.
        :param float low_energy: Threshold value for low energy highlighting.
        :param list[str] | None revisions: List of revisions to be included in arc flash scenario creation.
        :param args: Additional arguments for Worker initialization.
        :param kwargs: Additional keyword arguments for Worker initialization.
        """
        super().__init__(input_dir_path, output_dir_path, create_scenarios, run_scenarios, exclude_startswith,
                         exclude_contains, exclude_except, create_table, *args, **kwargs)
        self.use_si_units = use_si_units
        self.high_energy = high_energy
        self.low_energy = low_energy
        self.revisions = revisions
        self.scenario_class = lambda: ArcFlashScenario(url, revisions)

    def execute_data_parsing(self) -> None:
        """
        Executes the parsing of ANSI arc flash data by using the ArcFlashParser class.
        Parses data from the input directory and stores it in the instance variable.
        """
        af_parser = ArcFlashParser(self.input_dir_path)
        af_parser.extract_ansi_af_data()
        af_parser.parse_ansi_af_data(self.use_si_units, self.exclude_startswith,
                                     self.exclude_contains, self.exclude_except)
        self.parsed_ansi_data = af_parser.parsed_ansi_data

    def execute_data_export(self) -> Path:
        """
        Executes the export of parsed arc flash data to an Excel workbook.
        Creates headers, adds data, formats the sheet, and highlights high energy values.
        Saves the workbook to the output directory.

        :return: The path to the saved Excel workbook.
        :rtype: Path
        """
        af_exporter = ArcFlashExporter()
        af_exporter.create_headers(self.use_si_units)
        af_exporter.add_data(self.parsed_ansi_data)
        af_exporter.format_sheet(0, HEADER_ROW, len(AF_CONST_COLS), 0, WIDTH_COL_LRG)
        af_exporter.highlight_high_energy(self.low_energy, self.high_energy)
        project_number = self.input_dir_path.stem
        filename = f'{project_number}_{AF_FILENAME}'
        wb_path = Path(self.output_dir_path, filename)
        af_exporter.save_workbook(wb_path)
        return wb_path
