import traceback
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal


class Worker(QThread):
    """
    A Worker class that performs background tasks related to scenario execution, data parsing, and data export.
    Emits signals when an error occurs or when a process finishes.

    Attributes:
        error_occurred (pyqtSignal): Signal emitted when an error occurs, carrying the error message.
        process_finished (pyqtSignal): Signal emitted when a process finishes, carrying the output path.
    """
    error_occurred = pyqtSignal(str)
    process_finished = pyqtSignal(str)

    def __init__(self, input_dir_path: Path, output_dir_path: Path, create_scenarios: bool, run_scenarios: bool,
                 exclude_startswith: list, exclude_contains: list, create_table: bool, *args, **kwargs):
        """
        Initializes the Worker with required parameters for data processing tasks.

        :param Path input_dir_path: The directory path for input data.
        :param Path output_dir_path: The directory path for output data.
        :param bool create_scenarios: A flag to determine whether to create scenarios.
        :param bool run_scenarios: A flag to determine whether to run scenarios.
        :param list exclude_startswith: List of strings to exclude files that start with specified prefixes.
        :param list exclude_contains: List of strings to exclude files that contain specified substrings.
        :param bool create_table: A flag to determine whether to create an Excel table.
        """
        super().__init__(*args, **kwargs)
        self.input_dir_path = input_dir_path
        self.output_dir_path = output_dir_path
        self.create_scenarios = create_scenarios
        self.run_scenarios = run_scenarios
        self.exclude_startswith = exclude_startswith
        self.exclude_contains = exclude_contains
        self.create_table = create_table
        self.scenario_class = None
        self.parsed_ansi_data = None
        self.parsed_iec_data = None

    def execute_scenarios(self):
        """
        Executes the creation and/or running of scenarios based on the initialization parameters.
        Updates input and output directory paths as needed.
        """
        if self.create_scenarios:
            scenario = self.scenario_class()
            scenario.create_scenarios()
            self.input_dir_path = scenario.get_project_dir()
            if self.run_scenarios:
                scenario.run_scenarios()
            if str(self.output_dir_path) == '.':
                self.output_dir_path = self.input_dir_path

    def execute_data_parsing(self):
        """
        Parses data from input files. Abstract method for inheritance.
        """
        pass

    def execute_data_export(self):
        """
        Exports parsed data to the specified output directory. Abstract method for inheritance.
        """
        pass

    def start_next_process(self):
        """
        Starts the next process after the thread finishes. Abstract method for inheritance.
        """
        pass

    def run(self):
        """
        The main method executed when the thread starts. Handles the execution of scenarios,
        data parsing, and data export, emitting signals on completion or error.
        """
        try:
            self.execute_scenarios()
            output_path = None
            if self.create_table:
                self.execute_data_parsing()
                output_path = self.execute_data_export()
            self.process_finished.emit(str(output_path))
            self.start_next_process()

        except ConnectionError as ce:
            prompt = '. Please make sure ETAP Datahub is up and running.'
            self.error_occurred.emit(str(ce.args[0]) + prompt)
        except AttributeError as ae:
            prompt = '. Please make sure that the Device Duty files are up to date.'
            self.error_occurred.emit(str(ae.args[0]) + prompt)
        except PermissionError as pe:
            prompt = '. Please make sure that the file with the same name is closed.'
            self.error_occurred.emit(str(pe) + prompt)
        except Exception as e:
            print(traceback.format_exc())
            self.error_occurred.emit(str(e))
