import os
from pathlib import Path
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox
from worker.worker_af import ArcFlashWorker
from worker.worker_dd import DeviceDutyWorker
from .interface_ui import Ui_MainWindow


class Interface(QMainWindow, Ui_MainWindow):
    def __init__(self, app_path, *args, **kwargs):
        super(Interface, self).__init__(*args, **kwargs)
        self.icon_path = str(Path(app_path, './res', 'icon.ico'))
        self.setWindowIcon(QIcon(self.icon_path))
        self.app_path = None
        self.af_worker = None
        self.dd_worker = None
        self.etap_dir_conn = None
        self.setupUi(self)
        self.connect_buttons()
        self.add_connections()
        self.datahub_note.setVisible(False)
        self.exclude_revisions_input.setVisible(False)
        self.show()

    def show_file_browser_input(self):
        file_dialog = QFileDialog().getExistingDirectory()
        self.etap_dir.setText(file_dialog)

    def show_file_browser_output(self):
        file_dialog = QFileDialog().getExistingDirectory()
        self.output_dir.setText(file_dialog)

    def connect_buttons(self):
        self.browse_btn.clicked.connect(self.show_file_browser_input)
        self.browse_btn_2.clicked.connect(self.show_file_browser_output)
        self.generate_btn.clicked.connect(self.execute_worker_thread)
        self.clear_all_btn.clicked.connect(self.clear_all_inputs)

    def set_default_output_dir(self, is_checked):
        if is_checked:
            self.output_dir.setText(self.etap_dir.text())
            self.etap_dir_conn = self.etap_dir.textChanged.connect(self.output_dir.setText)
        else:
            self.output_dir.clear()
            self.etap_dir.textChanged.disconnect(self.etap_dir_conn)

    def execute_worker_thread(self):
        port = self.port.text()
        input_dir_path = Path(self.etap_dir.text())
        output_dir_path = Path(self.output_dir.text())
        create_scenarios = self.create_scenarios_checkbox.isChecked()
        run_scenarios = self.run_scenarios_checkbox.isChecked()
        exclude_startswith = self.split_tags(self.exclude_start_input.text())
        exclude_contains = self.split_tags(self.exclude_contain_input.text())
        create_table = self.create_reports_checkbox.isChecked()
        calculate_sw = self.sw_checkbox.isChecked()
        calculate_swgr = self.swgr_checkbox.isChecked()
        add_series_ratings = self.series_rating_checkbox.isChecked()
        mark_assumed = self.mark_assumed_checkbox.isChecked()
        high_energy = self.high_energy_box.value()
        low_energy = self.low_energy_box.value()

        arg_list_device_duty = [port, input_dir_path, output_dir_path, create_scenarios, run_scenarios,
                                exclude_startswith, exclude_contains, create_table, calculate_sw,
                                calculate_swgr, add_series_ratings, mark_assumed]

        arg_list_arc_flash = [port, input_dir_path, output_dir_path, create_scenarios, run_scenarios,
                              exclude_startswith, exclude_contains, create_table, high_energy, low_energy]

        if not self.validate_inputs():
            return

        if self.device_duty_checkbox.isChecked():
            self.dd_worker = DeviceDutyWorker(*arg_list_device_duty)
            self.dd_worker.error_occurred.connect(self.handle_error)
            self.dd_worker.process_finished.connect(self.handle_finished)
            run_arc_flash = lambda: self.run_arc_flash(*arg_list_arc_flash)
            self.dd_worker.start_arc_flash_process.connect(run_arc_flash)
            self.dd_worker.start()
        else:
            self.run_arc_flash(*arg_list_arc_flash)

    def run_arc_flash(self, *args):
        if self.arc_flash_checkbox.isChecked():
            self.af_worker = ArcFlashWorker(*args)
            self.af_worker.error_occurred.connect(self.handle_error)
            self.af_worker.process_finished.connect(self.handle_finished)
            self.af_worker.start()

    def handle_error(self, text):
        self.show_message('Runtime Error', text, QMessageBox.Critical)

    def show_message(self, title, message, icon=QMessageBox.Warning):
        message_box = QMessageBox()
        message_box.setWindowIcon(QIcon(self.icon_path))
        message_box.setIcon(icon)
        message_box.setText(message)
        message_box.setWindowTitle(title)
        message_box.setStandardButtons(QMessageBox.Ok)
        retval = message_box.exec_()

    @staticmethod
    def handle_finished(wb_path):
        if Path(wb_path).is_file():
            os.startfile(wb_path)

    def validate_inputs(self):
        port = self.port.text()
        input_dir_path = self.etap_dir.text()
        output_dir_path = self.output_dir.text()
        port_required = (self.create_scenarios_checkbox.isChecked() or
                         self.series_rating_checkbox.isChecked() or
                         self.mark_assumed_checkbox.isChecked())
        if not self.device_duty_checkbox.isChecked() and not self.arc_flash_checkbox.isChecked():
            self.show_message('Input Error', 'Please select a study.')
            return False
        if not self.create_reports_checkbox.isChecked() and not self.create_scenarios_checkbox.isChecked():
            self.show_message('Input Error', 'Please select an option to continue.')
            return False
        if port_required and (not port or not port.isnumeric() or len(port) < 5):
            self.show_message('Data Validation Error', 'Please enter a valid ETAP Datahub port number.')
            return False
        if self.etap_dir.isEnabled() and (not input_dir_path or not Path(input_dir_path).is_dir()):
            self.show_message('Data Validation Error', 'Please enter a valid ETAP project directory path.')
            return False
        if self.output_dir.isEnabled() and (not output_dir_path or not Path(output_dir_path).is_dir()):
            self.show_message('Data Validation Error', 'Please enter a valid output file directory path.')
            return False
        return True

    def handle_options_toggle(self, is_checked):
        options_group = [
            self.series_rating_checkbox,
            self.mark_assumed_checkbox,
            self.create_scenarios_checkbox,
            self.run_scenarios_checkbox
        ]
        if not is_checked and not any(option.isChecked() for option in options_group):
            self.datahub_note.setVisible(False)
        else:
            self.datahub_note.setVisible(True)

    def add_connections(self):
        self.create_scenarios_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.run_scenarios_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.mark_assumed_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.series_rating_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.etap_dir_checkbox.clicked['bool'].connect(self.set_default_output_dir)

    @staticmethod
    def split_tags(tags, delimiter=';'):
        return [tag.strip() for tag in filter(None, tags.split(delimiter))]

    def clear_all_inputs(self):
        checkboxes = [
            self.device_duty_checkbox,
            self.arc_flash_checkbox,
            self.create_scenarios_checkbox,
            self.run_scenarios_checkbox,
            self.mark_assumed_checkbox,
            self.series_rating_checkbox,
            self.sw_checkbox,
            self.swgr_checkbox,
            self.etap_dir_checkbox
        ]

        lines = [
            self.etap_dir,
            self.output_dir,
            self.exclude_start_input,
            self.exclude_contain_input
        ]

        for line in lines:
            line.clear()

        for checkbox in checkboxes:
            checkbox.setChecked(False)
