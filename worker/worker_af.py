from pathlib import Path
from consts.consts_af import FILE_NAME_SUFFIX
from controller.parser_af import ArcFlashParser
from exporter.exporter_af import ArcFlashExporter
from scenario.scenario_af import ArcFlashScenario
from worker.worker import Worker


class ArcFlashWorker(Worker):
    """
    ArcFlashWorker class for handling the creation, parsing, and exporting of arc flash data.
    Inherits from Worker class and overrides specific methods for arc flash data processing.
    """

    def __init__(self, port: int, input_dir_path: Path, output_dir_path: Path, create_scenarios: bool,
                 run_scenarios: bool, exclude_startswith: list[str], exclude_contains: list[str], create_table: bool,
                 high_energy: float, low_energy: float, *args, **kwargs):
        """
        Initializes the ArcFlashWorker with specific parameters for arc flash processing.

        :param int port: Port number for datahub connection.
        :param Path input_dir_path: Path to the input directory containing data.
        :param Path output_dir_path: Path to the output directory for saving results.
        :param bool create_scenarios: Flag to indicate if scenarios should be created.
        :param bool run_scenarios: Flag to indicate if scenarios should be run.
        :param list exclude_startswith: List of strings; files starting with these prefixes will be excluded.
        :param list exclude_contains: List of strings; files containing these substrings will be excluded.
        :param float high_energy: Threshold value for high energy highlighting.
        :param float low_energy: Threshold value for low energy highlighting.
        :param args: Additional arguments for Worker initialization.
        :param kwargs: Additional keyword arguments for Worker initialization.
        """
        super().__init__(port, input_dir_path, output_dir_path, create_scenarios, run_scenarios,
                         exclude_startswith, exclude_contains, create_table, *args, **kwargs)
        self.high_energy = high_energy
        self.low_energy = low_energy
        self.scenario_class = ArcFlashScenario

    def execute_data_parsing(self) -> None:
        """
        Executes the parsing of ANSI arc flash data by using the ArcFlashParser class.
        Parses data from the input directory and stores it in the instance variable.
        """
        af_parser = ArcFlashParser(self.input_dir_path)
        af_parser.extract_ansi_af_data()
        af_parser.parse_ansi_af_data(self.exclude_startswith, self.exclude_contains)
        self.parsed_ansi_data = af_parser.ansi_af_data

    def execute_data_export(self) -> Path:
        """
        Executes the export of parsed arc flash data to an Excel workbook.
        Creates headers, adds data, formats the sheet, and highlights high energy values.
        Saves the workbook to the output directory.

        :return: The path to the saved Excel workbook.
        :rtype: Path
        """
        af_exporter = ArcFlashExporter()
        af_exporter.create_headers()
        af_exporter.add_data(self.parsed_ansi_data)
        af_exporter.format_sheet()
        af_exporter.highlight_high_energy(self.low_energy, self.high_energy)
        project_number = self.input_dir_path.stem
        filename = f'{project_number}_{FILE_NAME_SUFFIX}'
        wb_path = Path(self.output_dir_path, filename)
        af_exporter.save_workbook(wb_path)
        return wb_path
