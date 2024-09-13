from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
from consts.consts_dd import mom_const_cols, mom_alt_cols, int_const_cols, int_alt_cols, spec_keys_iec_int
from consts.consts_dd import int_iec_alt_cols, spec_keys_mom, spec_keys_int
from exporter.exporter_dd import DeviceDutyExporter
from controller.parser_dd import DeviceDutyParser
from consts.consts_dd import MODES, ANSI_EXT, IEC_EXT
from requests.exceptions import ConnectionError
from scenario.scenario_dd import DeviceDutyScenario


class Worker(QThread):
    error_occurred = pyqtSignal(str)
    process_finished = pyqtSignal(str)

    def __init__(self, port, create_scenarios, run_scenarios, input_dir_path, output_dir_path,
                 calculate_sw, calculate_swgr, add_series_ratings, mark_assumed,
                 exclude_startswith, exclude_contains, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.datahub_port = port
        self.create_scenarios = create_scenarios
        self.run_scenarios = run_scenarios
        self.input_dir_path = input_dir_path
        self.output_dir_path = output_dir_path
        self.calculate_sw = calculate_sw
        self.calculate_swgr = calculate_swgr
        self.add_series_ratings = add_series_ratings
        self.mark_assumed = mark_assumed
        self.exclude_startswith = exclude_startswith
        self.exclude_contains = exclude_contains

    def run(self):
        try:
            if self.create_scenarios:
                dd_scenario = DeviceDutyScenario(self.datahub_port)
                dd_scenario.create_scenarios()
                self.input_dir_path = dd_scenario.get_project_dir()
                if self.run_scenarios:
                    dd_scenario.run_scenarios()
                if not self.output_dir_path:
                    self.output_dir_path = self.input_dir_path

            dd_parser = DeviceDutyParser(self.input_dir_path)
            dd_parser.extract_ansi_data()
            dd_parser.parse_ansi_data(self.exclude_startswith, self.exclude_contains,
                                      self.calculate_sw, self.calculate_swgr)
            dd_parser.extract_iec_data()
            dd_parser.parse_iec_data(self.exclude_startswith, self.exclude_contains)

            if self.add_series_ratings or self.mark_assumed:
                dd_parser.connect_to_etap(self.datahub_port)
                if self.add_series_ratings:
                    dd_parser.process_series_rated_equipment()
                if self.mark_assumed:
                    dd_parser.process_assumed_equipment()

            dd_exporter = DeviceDutyExporter()
            dd_exporter.set_ansi_data(dd_parser.parsed_ansi_data)
            dd_exporter.set_iec_data(dd_parser.parsed_iec_data)

            dd_exporter.create_headers(0, mom_const_cols, mom_alt_cols, 'Momentary Duty')
            dd_exporter.create_headers(1, int_const_cols, int_alt_cols, 'Interrupting Duty')
            dd_exporter.create_headers(2, int_const_cols, int_iec_alt_cols, 'Interrupting Duty')

            dd_exporter.insert_data(0, MODES['Momentary'], spec_keys_mom)
            dd_exporter.insert_data(1, MODES['Interrupt'], spec_keys_int)
            dd_exporter.insert_data(2, MODES['Interrupt'], spec_keys_iec_int, dataset='iec')

            dd_exporter.format_headers(0)
            dd_exporter.format_headers(1)
            dd_exporter.format_headers(2)

            dd_exporter.format_sheet(0, len(mom_const_cols), len(mom_alt_cols), 16)
            dd_exporter.format_sheet(1, len(int_const_cols), len(int_alt_cols), 22)
            dd_exporter.format_sheet(2, len(int_const_cols), len(int_iec_alt_cols), 16)

            project_number = self.input_dir_path.stem
            filename = f'{project_number}_Device Duty Report.xlsx'
            wb_path = Path(self.output_dir_path, filename)
            dd_exporter.save_workbook(wb_path)
            self.process_finished.emit(str(wb_path))

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

