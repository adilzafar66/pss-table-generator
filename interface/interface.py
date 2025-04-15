import os
import pickle
from pathlib import Path
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox

from consts.common import HTTPS, HTTP
from worker.worker_af import ArcFlashWorker
from worker.worker_dd import DeviceDutyWorker
from worker.worker_sc import ShortCircuitWorker
from .interface_ui import Ui_MainWindow


class Interface(QMainWindow, Ui_MainWindow):
    RUNTIME_ERROR_TITLE = 'Runtime Error'
    INPUT_ERROR_TITLE = 'Input Error'
    DATA_VALIDATION_ERROR_TITLE = 'Data Validation Error'
    RUNTIME_ERROR_MSG = 'Please select a study.'
    NO_OPTION_SELECTED_MSG = 'Please select an option to continue.'
    INVALID_PORT_MSG = 'Please enter a valid ETAP Datahub port number.'
    INVALID_ETAP_DIR_MSG = 'Please enter a valid ETAP project directory path.'
    INVALID_OUTPUT_DIR_MSG = 'Please enter a valid output file directory path.'
    PORT_FILE = 'port_number.pickle'

    def __init__(self, app_path: Path, *args, **kwargs):
        """
        Initializes the interface with the given application path and optional arguments.

        :param str app_path: Path to the application directory.
        :param args: Additional positional arguments.
        :param kwargs: Additional keyword arguments.
        """
        super(Interface, self).__init__(*args, **kwargs)
        self.icon_path = str(Path(app_path, './res', 'icon.ico'))
        self.default_path = Path.home() / 'Documents' / 'PTG Data'
        self.setWindowIcon(QIcon(self.icon_path))
        self.setupUi(self)
        self.set_logo(app_path)
        self.connect_buttons()
        self.add_connections()
        self.sc_worker = None
        self.af_worker = None
        self.dd_worker = None
        self.etap_dir_conn = None
        self.load_port_number(self.default_path / self.PORT_FILE)
        self.include_revisions_input.setVisible(False)
        self.datahub_note.setVisible(False)
        self.device_duty_group.setDisabled(True)
        self.arc_flash_layout.setDisabled(True)
        self.setMinimumWidth(475)
        self.adjustSize()
        self.show()

    def set_logo(self, app_path):
        logo_path = str(Path(app_path, './res', 'logo_title.png'))
        logo = QPixmap(logo_path)
        self.title.setPixmap(logo)

    def show_file_browser_input(self) -> None:
        """
        Opens a file dialog to select the ETAP project directory and sets the directory path.
        """
        file_dialog = QFileDialog().getExistingDirectory()
        self.etap_dir.setText(file_dialog)

    def show_file_browser_output(self) -> None:
        """
        Opens a file dialog to select the output directory and sets the directory path.
        """
        file_dialog = QFileDialog().getExistingDirectory()
        self.output_dir.setText(file_dialog)

    def connect_buttons(self) -> None:
        """
        Connects buttons in the UI to their respective event handlers.
        """
        self.browse_btn.clicked.connect(self.show_file_browser_input)
        self.browse_btn_2.clicked.connect(self.show_file_browser_output)
        self.generate_btn.clicked.connect(self.execute_worker_thread)
        self.clear_all_btn.clicked.connect(self.clear_all_inputs)

    def set_default_output_dir(self, is_checked: bool) -> None:
        """
        Sets the default output directory to the ETAP directory if checked, otherwise clears it.

        :param bool is_checked: Whether the option to set the default output directory is checked.
        """
        if is_checked:
            self.output_dir.setText(self.etap_dir.text())
            self.etap_dir_conn = self.etap_dir.textChanged.connect(self.output_dir.setText)
        else:
            self.output_dir.clear()
            self.etap_dir.textChanged.disconnect(self.etap_dir_conn)

    def get_datahub_url(self):
        protocol = HTTPS if self.radio_etap24.isChecked() else HTTP
        return f'{protocol}://localhost:{self.port.text()}'

    def get_analysis_args(self):
        url = self.get_datahub_url()
        input_dir_path = Path(self.etap_dir.text())
        output_dir_path = Path(self.output_dir.text())
        create_scenarios = self.create_scenarios_checkbox.isChecked()
        run_scenarios = self.run_scenarios_checkbox.isChecked()
        exclude_startswith = self.split_tags(self.exclude_start_input.text())
        exclude_contains = self.split_tags(self.exclude_contain_input.text())
        exclude_except = self.get_exclude_except()
        create_table = self.create_reports_checkbox.isChecked()
        add_switches = self.sw_checkbox.isChecked()
        use_all_sw_configs = self.use_all_checkbox.isChecked()
        add_series_ratings = self.series_rating_checkbox.isChecked()
        mark_assumed = self.mark_assumed_checkbox.isChecked()
        use_si_units = self.si_units_checkbox.isChecked()
        high_energy = self.high_energy_box.value()
        low_energy = self.low_energy_box.value()
        revisions = self.get_revisions()

        args_short_circuit = [url, input_dir_path,
                              output_dir_path, create_scenarios,
                              run_scenarios, exclude_startswith,
                              exclude_contains, exclude_except, create_table]

        args_device_duty = [url, input_dir_path, output_dir_path,
                            create_scenarios, run_scenarios, exclude_startswith,
                            exclude_contains, exclude_except, create_table,
                            add_switches, use_all_sw_configs, add_series_ratings, mark_assumed]

        args_arc_flash = [url, input_dir_path, output_dir_path,
                          create_scenarios, run_scenarios, exclude_startswith,
                          exclude_contains, exclude_except, create_table,
                          use_si_units, high_energy, low_energy, revisions]

        return args_short_circuit, args_device_duty, args_arc_flash

    def execute_worker_thread(self) -> None:
        """
        Validates inputs and executes Short Circuit, Device Duty and/or Arc Flash worker threads.
        """
        analysis_args = self.get_analysis_args()
        if not self.validate_inputs():
            return

        if self.short_circuit_checkbox.isChecked():
            self.run_short_circuit(*analysis_args)
            return

        if self.device_duty_checkbox.isChecked():
            self.run_device_duty(analysis_args[1], analysis_args[2])
            return

        if self.arc_flash_checkbox.isChecked():
            self.run_arc_flash(analysis_args[2])
            return

    def run_short_circuit(self, args_short_circuit: list, args_device_duty: list, args_arc_flash: list) -> None:
        """
        Runs the Short Circuit worker if the Short Circuit checkbox is checked.

        :param list args_short_circuit: Arguments to pass to the Short Circuit worker.
        :param list args_device_duty: Arguments to pass to the Device Duty worker.
        :param list args_arc_flash: Arguments to pass to the Arc Flash worker.
        """
        if self.short_circuit_checkbox.isChecked():
            self.sc_worker = ShortCircuitWorker(*args_short_circuit)
            self.sc_worker.error_occurred.connect(self.handle_error)
            run_device_duty = lambda: self.run_device_duty(args_device_duty, args_arc_flash)
            self.sc_worker.start_device_duty_process.connect(run_device_duty)
            self.sc_worker.start()

    def run_device_duty(self, args_device_duty: list, args_arc_flash: list) -> None:
        """
        Runs the Device Duty worker if the Device Duty checkbox is checked.

        :param list args_device_duty: Arguments to pass to the Device Duty worker.
        :param list args_arc_flash: Arguments to pass to the Arc Flash worker.
        """
        if self.device_duty_checkbox.isChecked():
            self.dd_worker = DeviceDutyWorker(*args_device_duty)
            self.dd_worker.error_occurred.connect(self.handle_error)
            self.dd_worker.process_finished.connect(self.handle_finished)
            run_arc_flash = lambda: self.run_arc_flash(args_arc_flash)
            self.dd_worker.start_arc_flash_process.connect(run_arc_flash)
            self.dd_worker.start()

    def run_arc_flash(self, args_arc_flash: list) -> None:
        """
        Runs the Arc Flash worker if the Arc Flash checkbox is checked.

        :param list args_arc_flash: Arguments to pass to the Arc Flash worker.
        """
        if self.arc_flash_checkbox.isChecked():
            self.af_worker = ArcFlashWorker(*args_arc_flash)
            self.af_worker.error_occurred.connect(self.handle_error)
            self.af_worker.process_finished.connect(self.handle_finished)
            self.af_worker.start()

    def handle_error(self, text: str) -> None:
        """
        Handles errors by displaying a message box with the given error message.

        :param str text: The error message to display.
        """
        self.show_message('Runtime Error', text, icon=QMessageBox.Critical)

    def show_message(self, title: str, message: str, description: str = '',
                     icon: QMessageBox.Icon = QMessageBox.Warning) -> None:
        """
        Displays a message box with the given title, message, and icon.

        :param str title: The title of the message box.
        :param str message: The message to display.
        :param str description: Additional description for the message box.
        :param QMessageBox.Icon icon: The icon to display in the message box.
        """
        message_box = QMessageBox()
        message_box.setWindowIcon(QIcon(self.icon_path))
        message_box.setIcon(icon)
        message_box.setText(message)
        message_box.setWindowTitle(title)
        message_box.setInformativeText(description)
        message_box.setStandardButtons(QMessageBox.Ok)
        message_box.exec_()

    @staticmethod
    def handle_finished(wb_path: str) -> None:
        """
        Handles the completion of a process by opening the file if it exists.

        :param str wb_path: The path to the workbook file to open.
        """
        if Path(wb_path).is_file():
            os.startfile(wb_path)

    def validate_inputs(self) -> bool:
        """
        Validates the input fields in the UI and displays appropriate error messages.

        :return bool: True if all inputs are valid, False otherwise.
        """
        port = self.port.text()
        input_dir_path = self.etap_dir.text()
        output_dir_path = self.output_dir.text()

        port_required = (self.create_scenarios_checkbox.isChecked() or
                         self.series_rating_checkbox.isChecked() or
                         self.mark_assumed_checkbox.isChecked())

        if not self.device_duty_checkbox.isChecked() and not self.arc_flash_checkbox.isChecked() and \
                not self.short_circuit_checkbox.isChecked():
            self.show_message(self.INPUT_ERROR_TITLE, self.RUNTIME_ERROR_MSG)
            return False
        if not self.create_reports_checkbox.isChecked() and not self.create_scenarios_checkbox.isChecked():
            self.show_message(self.INPUT_ERROR_TITLE, self.NO_OPTION_SELECTED_MSG)
            return False
        if port_required and (not port or not port.isnumeric() or len(port) < 5):
            self.show_message(self.DATA_VALIDATION_ERROR_TITLE, self.INVALID_PORT_MSG)
            return False
        if self.etap_dir.isEnabled() and (not input_dir_path or not Path(input_dir_path).is_dir()):
            self.show_message(self.DATA_VALIDATION_ERROR_TITLE, self.INVALID_ETAP_DIR_MSG)
            return False
        if self.output_dir.isEnabled() and (not output_dir_path or not Path(output_dir_path).is_dir()):
            self.show_message(self.DATA_VALIDATION_ERROR_TITLE, self.INVALID_OUTPUT_DIR_MSG)
            return False
        return True

    def handle_options_toggle(self, is_checked: bool) -> None:
        """
        Toggles the visibility of the DataHub note based on option selections.

        :param bool is_checked: Whether an option checkbox is checked.
        """
        options_group = [
            self.series_rating_checkbox,
            self.mark_assumed_checkbox,
            self.create_scenarios_checkbox,
            self.run_scenarios_checkbox
        ]
        if not is_checked and not any(option.isChecked() for option in options_group):
            # self.etap_dir.setDisabled(False)
            self.datahub_note.setVisible(False)
        else:
            # self.etap_dir.clear()
            # self.etap_dir.setDisabled(True)
            self.datahub_note.setVisible(True)

    def add_connections(self) -> None:
        """
        Adds connections for toggling options and setting default directories.
        """
        self.create_scenarios_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.run_scenarios_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.mark_assumed_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.series_rating_checkbox.toggled['bool'].connect(self.handle_options_toggle)
        self.etap_dir_checkbox.clicked['bool'].connect(self.set_default_output_dir)
        self.action_save_port.triggered.connect(lambda: self.save_port_number(self.default_path / self.PORT_FILE))
        self.action_load_port.triggered.connect(lambda: self.load_port_number(self.default_path / self.PORT_FILE))

    def save_port_number(self, save_file):
        try:
            save_file.parent.mkdir(exist_ok=True)
            with open(save_file, 'wb') as f:
                save_obj = self.port.text()
                pickle.dump(save_obj, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as ex:
            self.show_message('Save Failed', 'Unable to Save Port Number', str(ex))

    def load_port_number(self, load_file):
        if load_file.is_file():
            try:
                with open(load_file, 'rb') as f:
                    data_obj = pickle.load(f)
                    self.port.setText(data_obj)
            except Exception as ex:
                self.show_message('Load Failed', 'Unable to Load Port Number', str(ex))

    @staticmethod
    def split_tags(tags: str, delimiter: str = ';') -> list:
        """
        Splits tags based on the given delimiter and strips whitespace.

        :param str tags: The string of tags to split.
        :param str delimiter: The delimiter used to split the tags.
        :return list: A list of cleaned tag strings.
        """
        return [tag.strip() for tag in filter(None, tags.split(delimiter))]

    def get_revisions(self) -> list | None:
        """
        Retrieves the revision information based on user selections.

        :return list | None: A list of revisions or None if all revisions are included.
        """
        if self.include_all_radio.isChecked():
            return None
        if self.include_base_radio.isChecked():
            return []
        if self.include_only_radio.isChecked():
            revisions = self.include_revisions_input.text()
            return self.split_tags(revisions)

    def get_exclude_except(self) -> list:
        is_exclude_except_checked = self.exclude_except_radio.isChecked()
        if is_exclude_except_checked:
            return self.split_tags(self.exclude_except_input.text())
        return []

    def clear_all_inputs(self) -> None:
        """
        Clears all input fields and resets checkboxes to their default state.
        """
        checkboxes = [
            self.device_duty_checkbox,
            self.arc_flash_checkbox,
            self.short_circuit_checkbox,
            self.create_scenarios_checkbox,
            self.run_scenarios_checkbox,
            self.create_reports_checkbox,
            self.mark_assumed_checkbox,
            self.series_rating_checkbox,
            self.sw_checkbox,
            self.use_all_checkbox,
            self.etap_dir_checkbox,
            self.si_units_checkbox
        ]
        lines = [
            self.etap_dir,
            self.output_dir,
            self.exclude_start_input,
            self.exclude_contain_input,
            self.exclude_except_input
        ]
        self.include_base_radio.setChecked(True)
        self.exclude_all_radio.setChecked(True)
        for line in lines:
            line.clear()
        for checkbox in checkboxes:
            checkbox.setChecked(False)
