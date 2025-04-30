from pathlib import Path
from consts import errors
from consts.errors import RUNTIME_ERROR_TITLE
from interface import utils
from PyQt5.QtGui import QIcon, QPixmap
from interface.interface_ui import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox
from consts.common import HTTPS, HTTP, PROGRAM_TITLE, PORT_FILE
from worker.worker_af import ArcFlashWorker
from worker.worker_dd import DeviceDutyWorker
from worker.worker_sc import ShortCircuitWorker


class Interface(QMainWindow, Ui_MainWindow):
    """
    Main application interface handling the UI interactions and worker thread management.
    """

    def __init__(self, app_path: Path, *args, **kwargs):
        """
        Initializes the interface window and sets up the UI.

        :param Path app_path: Path to the application root.
        :param args: Positional arguments.
        :param kwargs: Keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.icon_path = str(app_path / 'res' / 'icon.ico')
        self.default_path = Path.home() / 'Documents' / PROGRAM_TITLE
        self.setWindowIcon(QIcon(self.icon_path))
        self.setupUi(self)

        self.sc_worker = None
        self.dd_worker = None
        self.af_worker = None
        self.etap_dir_conn = None

        self._initialize_interface(app_path)

    def _initialize_interface(self, app_path: Path) -> None:
        """
        Initializes UI components and connects signals.

        :param Path app_path: Path to the application root.
        """
        self._set_logo(app_path)
        self._connect_buttons()
        self._add_additional_connections()
        self._load_port_settings()
        self._configure_initial_visibility()
        self.setMinimumWidth(475)
        self.adjustSize()
        self.show()

    def _set_logo(self, app_path: Path) -> None:
        """
        Sets the application logo in the interface.

        :param Path app_path: Path to the application root.
        """
        logo_path = str(app_path / 'res' / 'logo_title.png')
        self.title.setPixmap(QPixmap(logo_path))

    def _configure_initial_visibility(self) -> None:
        """
        Configures the initial visibility and disabled state of UI components.
        """
        self.include_revisions_input.setVisible(False)
        self.datahub_note.setVisible(False)
        self.device_duty_group.setDisabled(True)
        self.arc_flash_layout.setDisabled(True)

    def _connect_buttons(self) -> None:
        """
        Connects UI buttons to their respective slot methods.
        """
        self.browse_btn.clicked.connect(self._browse_input_directory)
        self.browse_btn_2.clicked.connect(self._browse_output_directory)
        self.generate_btn.clicked.connect(self._start_analysis)
        self.clear_all_btn.clicked.connect(self._clear_inputs)

    def _add_additional_connections(self) -> None:
        """
        Adds signal connections for toggling options and port actions.
        """
        option_checkboxes = [
            self.create_scenarios_checkbox,
            self.run_scenarios_checkbox,
            self.mark_assumed_checkbox,
            self.series_rating_checkbox
        ]
        for checkbox in option_checkboxes:
            checkbox.toggled.connect(self._toggle_datahub_note_visibility)

        self.etap_dir_checkbox.clicked.connect(self._sync_output_directory)
        self.action_save_port.triggered.connect(self._save_port_settings)
        self.action_load_port.triggered.connect(self._load_port_settings)

    def _browse_input_directory(self) -> None:
        """
        Opens a dialog to select an input directory.
        """
        selected_dir = QFileDialog.getExistingDirectory()
        if selected_dir:
            self.etap_dir.setText(selected_dir)

    def _browse_output_directory(self) -> None:
        """
        Opens a dialog to select an output directory.
        """
        selected_dir = QFileDialog.getExistingDirectory()
        if selected_dir:
            self.output_dir.setText(selected_dir)

    def _sync_output_directory(self, checked: bool) -> None:
        """
        Synchronizes output directory with input directory if checked.

        :param bool checked: Whether synchronization is enabled.
        """
        if checked:
            self.output_dir.setText(self.etap_dir.text())
            self.etap_dir_conn = self.etap_dir.textChanged.connect(self.output_dir.setText)
        else:
            self.output_dir.clear()
            if self.etap_dir_conn:
                self.etap_dir.textChanged.disconnect(self.etap_dir_conn)

    def _toggle_datahub_note_visibility(self, checked: bool) -> None:
        """
        Toggles visibility of the DataHub note based on option selections.

        :param bool checked: Checkbox state.
        """
        options = [
            self.create_scenarios_checkbox.isChecked(),
            self.run_scenarios_checkbox.isChecked(),
            self.series_rating_checkbox.isChecked(),
            self.mark_assumed_checkbox.isChecked()
        ]
        self.datahub_note.setVisible(any(options))

    def _save_port_settings(self) -> None:
        """
        Saves port settings to file.
        """
        port_file = self.default_path / PORT_FILE
        error_func = lambda ex: self._show_message(errors.SAVE_ERROR_TITLE, errors.PORT_SAVE_ERROR, str(ex))
        utils.save_line_edit_text(port_file, self.port, error_func)

    def _load_port_settings(self) -> None:
        """
        Loads port settings from file.
        """
        port_file = self.default_path / PORT_FILE
        error_func = lambda ex: self._show_message(errors.LOAD_ERROR_TITLE, errors.PORT_LOAD_ERROR, str(ex))
        utils.load_line_edit_text(port_file, self.port, error_func)

    def _start_analysis(self) -> None:
        """
        Initiates the analysis process based on user selections.
        """
        if not self._validate_inputs():
            return

        args_sc, args_dd, args_af = self._collect_analysis_arguments()

        if self.short_circuit_checkbox.isChecked():
            self._run_short_circuit(args_sc, args_dd, args_af)
        elif self.device_duty_checkbox.isChecked():
            self._run_device_duty(args_dd, args_af)
        elif self.arc_flash_checkbox.isChecked():
            self._run_arc_flash(args_af)

    def _collect_analysis_arguments(self) -> tuple:
        """
        Collects and structures the arguments required for each worker.

        :return tuple: Arguments for Short Circuit, Device Duty, and Arc Flash workers.
        """
        url = self._get_datahub_url()
        input_dir = Path(self.etap_dir.text())
        output_dir = Path(self.output_dir.text())

        common_args = [
            url, input_dir, output_dir,
            self.create_scenarios_checkbox.isChecked(),
            self.run_scenarios_checkbox.isChecked(),
            utils.split_string_tags(self.exclude_start_input.text()),
            utils.split_string_tags(self.exclude_contain_input.text()),
            self._get_exclude_except(),
            self.create_reports_checkbox.isChecked()
        ]

        device_duty_args = common_args + [
            self.sw_checkbox.isChecked(),
            self.use_all_checkbox.isChecked(),
            self.series_rating_checkbox.isChecked(),
            self.mark_assumed_checkbox.isChecked()
        ]

        arc_flash_args = common_args + [
            self.si_units_checkbox.isChecked(),
            self.high_energy_box.value(),
            self.low_energy_box.value(),
            self._get_revisions()
        ]

        return common_args, device_duty_args, arc_flash_args

    def _get_datahub_url(self) -> str:
        """
        Constructs the DataHub URL based on user selection.

        :return str: The constructed URL.
        """
        protocol = HTTPS if self.radio_etap24.isChecked() else HTTP
        return f'{protocol}://localhost:{self.port.text()}'

    def _get_revisions(self) -> list | None:
        """
        Retrieves revisions based on user selection.

        :return list | None: List of revisions or None.
        """
        if self.include_all_radio.isChecked():
            return None
        if self.include_base_radio.isChecked():
            return []
        return utils.split_string_tags(self.include_revisions_input.text())

    def _get_exclude_except(self) -> list:
        """
        Retrieves tags for the "exclude except" field.

        :return list: Tags entered for exclusion exception.
        """
        return utils.split_string_tags(
            self.exclude_except_input.text()) if self.exclude_except_radio.isChecked() else []

    def _run_short_circuit(self, sc_args: list, dd_args: list, af_args: list) -> None:
        """
        Runs the Short Circuit analysis worker.

        :param list sc_args: Short Circuit arguments.
        :param list dd_args: Device Duty arguments.
        :param list af_args: Arc Flash arguments.
        """
        self.sc_worker = ShortCircuitWorker(*sc_args)
        self.sc_worker.error_occurred.connect(self._handle_error)
        self.sc_worker.start_device_duty_process.connect(lambda: self._run_device_duty(dd_args, af_args))
        self.sc_worker.start()

    def _run_device_duty(self, dd_args: list, af_args: list) -> None:
        """
        Runs the Device Duty analysis worker.

        :param list dd_args: Device Duty arguments.
        :param list af_args: Arc Flash arguments.
        """
        if self.device_duty_checkbox.isChecked():
            self.dd_worker = DeviceDutyWorker(*dd_args)
            self.dd_worker.error_occurred.connect(self._handle_error)
            self.dd_worker.process_finished.connect(self._handle_process_finished)
            self.dd_worker.start_arc_flash_process.connect(lambda: self._run_arc_flash(af_args))
            self.dd_worker.start()

    def _run_arc_flash(self, af_args: list) -> None:
        """
        Runs the Arc Flash analysis worker.

        :param list af_args: Arc Flash arguments.
        """
        if self.arc_flash_checkbox.isChecked():
            self.af_worker = ArcFlashWorker(*af_args)
            self.af_worker.error_occurred.connect(self._handle_error)
            self.af_worker.process_finished.connect(self._handle_process_finished)
            self.af_worker.start()

    def _handle_error(self, message: str) -> None:
        """
        Handles runtime errors by showing a message.

        :param str message: The error message.
        """
        self._show_message(RUNTIME_ERROR_TITLE, message, icon=QMessageBox.Critical)

    def _show_message(self, title: str, message: str, description: str = '',
                      icon: QMessageBox.Icon = QMessageBox.Warning) -> None:
        """
        Displays a custom message box.

        :param str title: Title of the message box.
        :param str message: Main message content.
        :param str description: Optional additional description.
        :param QMessageBox.Icon icon: Message icon.
        """
        box = QMessageBox(self)
        box.setWindowIcon(QIcon(self.icon_path))
        box.setIcon(icon)
        box.setWindowTitle(title)
        box.setText(message)
        box.setInformativeText(description)
        box.setStandardButtons(QMessageBox.Ok)
        box.exec_()

    @staticmethod
    def _handle_process_finished(output_path: str) -> None:
        """
        Handles worker completion.

        :param str output_path: Path to the output file.
        """
        utils.open_file(Path(output_path))

    def _validate_inputs(self) -> bool:
        """
        Validates form fields before starting processes.

        :return bool: True if all validations pass, otherwise False.
        """
        if not any([self.device_duty_checkbox.isChecked(), self.arc_flash_checkbox.isChecked(),
                    self.short_circuit_checkbox.isChecked()]):
            self._show_message(errors.INPUT_ERROR_TITLE, errors.RUNTIME_ERROR_MSG)
            return False

        if not any([self.create_reports_checkbox.isChecked(), self.create_scenarios_checkbox.isChecked()]):
            self._show_message(errors.INPUT_ERROR_TITLE, errors.NO_OPTION_SELECTED_MSG)
            return False

        if any([
            self.create_scenarios_checkbox.isChecked(),
            self.series_rating_checkbox.isChecked(),
            self.mark_assumed_checkbox.isChecked()
        ]) and (not self.port.text().isnumeric() or len(self.port.text()) < 5):
            self._show_message(errors.DATA_VALIDATION_ERROR_TITLE, errors.INVALID_PORT_MSG)
            return False

        if self.etap_dir.isEnabled() and not Path(self.etap_dir.text()).is_dir():
            self._show_message(errors.DATA_VALIDATION_ERROR_TITLE, errors.INVALID_ETAP_DIR_MSG)
            return False

        if self.output_dir.isEnabled() and not Path(self.output_dir.text()).is_dir():
            self._show_message(errors.DATA_VALIDATION_ERROR_TITLE, errors.INVALID_OUTPUT_DIR_MSG)
            return False

        return True

    def _clear_inputs(self) -> None:
        """
        Clears all user inputs and resets UI state.
        """
        fields_to_clear = [
            self.etap_dir, self.output_dir,
            self.exclude_start_input, self.exclude_contain_input,
            self.exclude_except_input
        ]
        checkboxes_to_uncheck = [
            self.device_duty_checkbox, self.arc_flash_checkbox, self.short_circuit_checkbox,
            self.create_scenarios_checkbox, self.run_scenarios_checkbox, self.create_reports_checkbox,
            self.mark_assumed_checkbox, self.series_rating_checkbox, self.sw_checkbox,
            self.use_all_checkbox, self.etap_dir_checkbox, self.si_units_checkbox
        ]

        for field in fields_to_clear:
            field.clear()

        for checkbox in checkboxes_to_uncheck:
            checkbox.setChecked(False)

        self.include_base_radio.setChecked(True)
        self.exclude_all_radio.setChecked(True)
