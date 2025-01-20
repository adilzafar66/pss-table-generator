from pathlib import Path
from PyQt5.QtCore import pyqtSignal
from scenario.scenario_sc import ShortCircuitScenario
from worker.worker import Worker


class ShortCircuitWorker(Worker):
    """
    ShortCircuitWorker class for handling short circuit analysis, parsing, and exporting.
    Inherits from the Worker class and provides specialized methods for short circuit operations.
    """

    start_device_duty_process = pyqtSignal()

    def __init__(self, url: str, input_dir_path: Path, output_dir_path: Path, create_scenarios: bool,
                 run_scenarios: bool, exclude_startswith: list[str], exclude_contains: list[str],
                 exclude_except: list[str], create_table: bool, *args, **kwargs):
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
        :param args: Additional arguments for Worker initialization.
        :param kwargs: Additional keyword arguments for Worker initialization.
        """
        super().__init__(input_dir_path, output_dir_path, create_scenarios, run_scenarios, exclude_startswith,
                         exclude_contains, exclude_except, create_table, *args, **kwargs)
        self.scenario_class = lambda: ShortCircuitScenario(url)

    def start_next_process(self):
        """
        Sends the trigger to start the device duty process after the thread finishes.
        """
        self.start_device_duty_process.emit()
