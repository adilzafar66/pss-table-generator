from pathlib import Path
from PyQt5.QtCore import Qt, QTimer
from consts import errors
from interface import utils
from PyQt5.QtGui import QIcon, QPixmap
from interface.interface_pb import Ui_Dialog
from interface.interface_ui import Ui_MainWindow
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox, QDialog
from consts.common import PROGRAM_TITLE, DEFAULTS_FILENAME, JSON_FORMAT
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
        self._configure_initial_visibility()
        self.load_default_inputs()
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
        self.arc_flash_group.setDisabled(True)

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

        self.etap_dir_checkbox.toggled.connect(self._sync_output_directory)
        self.action_save_defaults.triggered.connect(self.save_default_inputs)
        self.action_open_defaults.triggered.connect(self.load_default_inputs)
        self.action_save.triggered.connect(self.save_as_inputs)
        self.action_open.triggered.connect(self.load_as_inputs)

        self.action_exclude.triggered.connect(lambda a: self.set_visible(self.exclude_group, a))
        self.action_device_duty.triggered.connect(lambda a: self.set_visible(self.device_duty_group, a))
        self.action_arc_flash.triggered.connect(lambda a: self.set_visible(self.arc_flash_group, a))
        self.include_only_radio.toggled.connect(lambda a: self.set_visible(self.include_revisions_input, a))

    def set_visible(self, element, visible):
        element.setVisible(visible)
        QTimer.singleShot(0, self.adjustSize)

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
        self.set_visible(self.datahub_note, any(options))

    def _start_analysis(self) -> None:
        """
        Initiates the analysis process based on user selections.
        """
        if not self._validate_inputs():
            return

        self._show_progress_dialog()

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

        short_circuit_args = common_args + [
            self.use_all_checkbox.isChecked(),
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

        return short_circuit_args, device_duty_args, arc_flash_args

    def _get_datahub_url(self) -> str:
        """
        Constructs the DataHub URL based on user selection.

        :return str: The constructed URL.
        """
        protocol, port_number = utils.get_datahub_info(self.etap_dir.text())
        return f'{protocol}://localhost:{port_number}'

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
        self.sc_worker.process_finished.connect(self._handle_process_finished)
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
        self.progress_dialog and self.progress_dialog.close()
        self._show_message(errors.RUNTIME_ERROR_TITLE, message, icon=QMessageBox.Critical)

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

    def _show_progress_dialog(self) -> None:
        """
        Creates and displays a progress dialog until the process is completed.
        """
        self.progress_dialog = Dialog(QIcon(self.icon_path))
        self.progress_dialog.setWindowModality(Qt.ApplicationModal)
        self.progress_dialog.show()

    def _handle_process_finished(self, output_path: str) -> None:
        """
        Handles worker completion.

        :param str output_path: Path to the output file.
        """
        utils.open_file(Path(output_path))
        threads = [worker for worker in [self.sc_worker, self.dd_worker, self.af_worker] if worker is not None]
        thread_cbs = [self.short_circuit_checkbox, self.device_duty_checkbox, self.arc_flash_checkbox]
        checked_cbs = [cb for cb in thread_cbs if cb.isChecked()]

        # Check if all threads are finished and the counts match
        if all(thread.isFinished() for thread in threads) and len(threads) == len(checked_cbs):
            self.progress_dialog.close()
            self.sc_worker = None
            self.dd_worker = None
            self.af_worker = None

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

        def is_input_empty(i):
            return not i.text() or not Path(i.text()).is_dir()

        if is_input_empty(self.etap_dir):
            self._show_message(errors.DATA_VALIDATION_ERROR_TITLE, errors.INVALID_ETAP_DIR_MSG)
            return False

        if self.output_dir.isEnabled() and is_input_empty(self.output_dir):
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

    def save_default_inputs(self) -> None:
        """
        Saves current input values to the default file.
        """
        utils.save_inputs(self, self.default_path / DEFAULTS_FILENAME)

    def load_default_inputs(self) -> None:
        """
        Loads input values from the default file.
        """
        utils.load_inputs(self, self.default_path / DEFAULTS_FILENAME)

    def save_as_inputs(self) -> None:
        """
        Opens a file dialog for the user to choose a file location and
        saves the current input values to that file.
        """
        filepath, _ = QFileDialog().getSaveFileName(self, 'Save', str(Path.home() / 'inputs.json'), JSON_FORMAT)
        filepath and utils.save_inputs(self, filepath)

    def load_as_inputs(self) -> None:
        """
        Opens a file dialog for the user to choose a file and
        loads input values from the selected file.
        """
        filepath, _ = QFileDialog().getOpenFileName(self, 'Open', str(Path.home()), JSON_FORMAT)
        filepath and utils.load_inputs(self, filepath)


class Dialog(QDialog, Ui_Dialog):
    """
    Custom dialog for showing progress during long-running tasks.
    """

    def __init__(self, icon: QIcon, *args, **kwargs):
        """
        Initializes the progress dialog with the given icon.

        :param QIcon icon: Icon for the dialog.
        :param args: Additional positional arguments.
        :param kwargs: Additional keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.setupUi(self)
        self.setWindowIcon(icon)
