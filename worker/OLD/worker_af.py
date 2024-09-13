from PyQt5.QtCore import QThread, pyqtSignal
from controller.parser_af import ArcFlashParser
from exporter.exporter_af import ArcFlashExporter
from scenario.scenario_af import ArcFlashScenario


class Worker(QThread):
    error_occurred = pyqtSignal(str)
    process_finished = pyqtSignal(str)

    def __init__(self, port, input_dir_path, output_dir_path, create_scenarios,
                 run_scenarios, high_energy, low_energy, exclude_startswith,
                 exclude_contains, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.datahub_port = port
        self.input_dir_path = input_dir_path
        self.output_dir_path = output_dir_path
        self.create_scenarios = create_scenarios
        self.run_scenarios = run_scenarios
        self.high_energy = high_energy
        self.low_energy = low_energy
        self.exclude_startswith = exclude_startswith
        self.exclude_contains = exclude_contains

    def run(self):
        try:
            if self.create_scenarios:
                af_scenario = ArcFlashScenario(self.datahub_port)
                af_scenario.create_scenarios()
                self.input_dir_path = af_scenario.get_project_dir()
                if self.run_scenarios:
                    af_scenario.run_scenarios()
                if not self.output_dir_path:
                    self.output_dir_path = self.input_dir_path

            af_parser = ArcFlashParser(self.input_dir_path)
            af_parser.extract_ansi_af_data()
            af_parser.parse_ansi_af_data(self.exclude_startswith, self.exclude_contains)

            af_exporter = ArcFlashExporter()
            af_exporter.create_headers()
            af_exporter.add_data(af_parser.ansi_af_data)
            af_exporter.format_sheet()
            af_exporter.highlight_high_energy(self.low_energy, self.high_energy)
            af_exporter.save_workbook()
            self.process_finished.emit()

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
            self.error_occurred.emit(str(e))


