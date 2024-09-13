# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '.\res\pss_report_gen_ui.ui'
#
# Created by: PyQt5 UI code generator 5.15.11
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowModality(QtCore.Qt.NonModal)
        MainWindow.resize(448, 783)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        MainWindow.setAnimated(True)
        MainWindow.setDocumentMode(False)
        MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout(self.centralwidget)
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.main_layout = QtWidgets.QVBoxLayout()
        self.main_layout.setContentsMargins(12, 0, 12, 0)
        self.main_layout.setObjectName("main_layout")
        self.title_layout = QtWidgets.QVBoxLayout()
        self.title_layout.setObjectName("title_layout")
        self.title = QtWidgets.QLabel(self.centralwidget)
        self.title.setObjectName("title")
        self.title_layout.addWidget(self.title)
        self.main_layout.addLayout(self.title_layout)
        self.port_group = QtWidgets.QGroupBox(self.centralwidget)
        self.port_group.setObjectName("port_group")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.port_group)
        self.verticalLayout.setObjectName("verticalLayout")
        self.port_layout = QtWidgets.QHBoxLayout()
        self.port_layout.setObjectName("port_layout")
        self.port = QtWidgets.QLineEdit(self.port_group)
        self.port.setMaximumSize(QtCore.QSize(60, 24))
        self.port.setMaxLength(5)
        self.port.setAlignment(QtCore.Qt.AlignCenter)
        self.port.setObjectName("port")
        self.port_layout.addWidget(self.port)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.port_layout.addItem(spacerItem)
        self.verticalLayout.addLayout(self.port_layout)
        self.main_layout.addWidget(self.port_group)
        self.report_options_grid = QtWidgets.QGridLayout()
        self.report_options_grid.setObjectName("report_options_grid")
        self.options_group = QtWidgets.QGroupBox(self.centralwidget)
        self.options_group.setObjectName("options_group")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout(self.options_group)
        self.verticalLayout_3.setSpacing(6)
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.options_grid = QtWidgets.QGridLayout()
        self.options_grid.setObjectName("options_grid")
        self.create_scenarios_checkbox = QtWidgets.QCheckBox(self.options_group)
        self.create_scenarios_checkbox.setChecked(False)
        self.create_scenarios_checkbox.setObjectName("create_scenarios_checkbox")
        self.options_grid.addWidget(self.create_scenarios_checkbox, 0, 0, 1, 1)
        self.run_scenarios_checkbox = QtWidgets.QCheckBox(self.options_group)
        self.run_scenarios_checkbox.setEnabled(False)
        self.run_scenarios_checkbox.setChecked(False)
        self.run_scenarios_checkbox.setObjectName("run_scenarios_checkbox")
        self.options_grid.addWidget(self.run_scenarios_checkbox, 1, 0, 1, 1)
        self.verticalLayout_3.addLayout(self.options_grid)
        self.report_options_grid.addWidget(self.options_group, 0, 1, 1, 1)
        self.report_group = QtWidgets.QGroupBox(self.centralwidget)
        self.report_group.setObjectName("report_group")
        self.verticalLayout_5 = QtWidgets.QVBoxLayout(self.report_group)
        self.verticalLayout_5.setObjectName("verticalLayout_5")
        self.report_grid = QtWidgets.QGridLayout()
        self.report_grid.setObjectName("report_grid")
        self.arc_flash_checkbox = QtWidgets.QCheckBox(self.report_group)
        self.arc_flash_checkbox.setChecked(True)
        self.arc_flash_checkbox.setObjectName("arc_flash_checkbox")
        self.report_grid.addWidget(self.arc_flash_checkbox, 1, 0, 1, 1)
        self.device_duty_checkbox = QtWidgets.QCheckBox(self.report_group)
        self.device_duty_checkbox.setChecked(True)
        self.device_duty_checkbox.setObjectName("device_duty_checkbox")
        self.report_grid.addWidget(self.device_duty_checkbox, 0, 0, 1, 1)
        self.verticalLayout_5.addLayout(self.report_grid)
        self.report_options_grid.addWidget(self.report_group, 0, 0, 1, 1)
        self.main_layout.addLayout(self.report_options_grid)
        self.device_duty_layout = QtWidgets.QVBoxLayout()
        self.device_duty_layout.setObjectName("device_duty_layout")
        self.device_duty_group = QtWidgets.QGroupBox(self.centralwidget)
        self.device_duty_group.setObjectName("device_duty_group")
        self.verticalLayout_10 = QtWidgets.QVBoxLayout(self.device_duty_group)
        self.verticalLayout_10.setObjectName("verticalLayout_10")
        self.device_duty_grid = QtWidgets.QGridLayout()
        self.device_duty_grid.setObjectName("device_duty_grid")
        self.sw_checkbox = QtWidgets.QCheckBox(self.device_duty_group)
        self.sw_checkbox.setChecked(True)
        self.sw_checkbox.setObjectName("sw_checkbox")
        self.device_duty_grid.addWidget(self.sw_checkbox, 1, 0, 1, 1)
        self.swgr_checkbox = QtWidgets.QCheckBox(self.device_duty_group)
        self.swgr_checkbox.setChecked(True)
        self.swgr_checkbox.setObjectName("swgr_checkbox")
        self.device_duty_grid.addWidget(self.swgr_checkbox, 1, 1, 1, 1)
        self.mark_assumed_checkbox = QtWidgets.QCheckBox(self.device_duty_group)
        self.mark_assumed_checkbox.setChecked(False)
        self.mark_assumed_checkbox.setObjectName("mark_assumed_checkbox")
        self.device_duty_grid.addWidget(self.mark_assumed_checkbox, 0, 1, 1, 1)
        self.series_rating_checkbox = QtWidgets.QCheckBox(self.device_duty_group)
        self.series_rating_checkbox.setChecked(False)
        self.series_rating_checkbox.setObjectName("series_rating_checkbox")
        self.device_duty_grid.addWidget(self.series_rating_checkbox, 0, 0, 1, 1)
        self.verticalLayout_10.addLayout(self.device_duty_grid)
        self.device_duty_layout.addWidget(self.device_duty_group)
        self.main_layout.addLayout(self.device_duty_layout)
        self.arc_flash_layout = QtWidgets.QGroupBox(self.centralwidget)
        self.arc_flash_layout.setObjectName("arc_flash_layout")
        self.verticalLayout_11 = QtWidgets.QVBoxLayout(self.arc_flash_layout)
        self.verticalLayout_11.setObjectName("verticalLayout_11")
        self.arc_flash_grid = QtWidgets.QGridLayout()
        self.arc_flash_grid.setObjectName("arc_flash_grid")
        self.nm_config_label = QtWidgets.QLabel(self.arc_flash_layout)
        self.nm_config_label.setObjectName("nm_config_label")
        self.arc_flash_grid.addWidget(self.nm_config_label, 0, 1, 1, 1)
        self.nm_config_layout = QtWidgets.QHBoxLayout()
        self.nm_config_layout.setObjectName("nm_config_layout")
        self.nm_switching_radio = QtWidgets.QRadioButton(self.arc_flash_layout)
        self.nm_switching_radio.setEnabled(False)
        self.nm_switching_radio.setChecked(True)
        self.nm_switching_radio.setObjectName("nm_switching_radio")
        self.nm_config_layout.addWidget(self.nm_switching_radio)
        self.nm_revisions_radio = QtWidgets.QRadioButton(self.arc_flash_layout)
        self.nm_revisions_radio.setEnabled(False)
        self.nm_revisions_radio.setChecked(False)
        self.nm_revisions_radio.setObjectName("nm_revisions_radio")
        self.nm_config_layout.addWidget(self.nm_revisions_radio)
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.nm_config_layout.addItem(spacerItem1)
        self.arc_flash_grid.addLayout(self.nm_config_layout, 0, 2, 1, 1)
        self.incident_energy_label = QtWidgets.QLabel(self.arc_flash_layout)
        self.incident_energy_label.setObjectName("incident_energy_label")
        self.arc_flash_grid.addWidget(self.incident_energy_label, 1, 1, 1, 1)
        self.incident_energy_layout = QtWidgets.QHBoxLayout()
        self.incident_energy_layout.setObjectName("incident_energy_layout")
        self.low_energy_layout = QtWidgets.QHBoxLayout()
        self.low_energy_layout.setObjectName("low_energy_layout")
        self.low_energy_label = QtWidgets.QLabel(self.arc_flash_layout)
        self.low_energy_label.setObjectName("low_energy_label")
        self.low_energy_layout.addWidget(self.low_energy_label)
        self.low_energy_box = QtWidgets.QDoubleSpinBox(self.arc_flash_layout)
        self.low_energy_box.setProperty("value", 12.0)
        self.low_energy_box.setObjectName("low_energy_box")
        self.low_energy_layout.addWidget(self.low_energy_box)
        self.incident_energy_layout.addLayout(self.low_energy_layout)
        self.high_energy_layout = QtWidgets.QHBoxLayout()
        self.high_energy_layout.setObjectName("high_energy_layout")
        self.high_energy_label = QtWidgets.QLabel(self.arc_flash_layout)
        self.high_energy_label.setObjectName("high_energy_label")
        self.high_energy_layout.addWidget(self.high_energy_label)
        self.high_energy_box = QtWidgets.QDoubleSpinBox(self.arc_flash_layout)
        self.high_energy_box.setProperty("value", 40.0)
        self.high_energy_box.setObjectName("high_energy_box")
        self.high_energy_layout.addWidget(self.high_energy_box)
        self.incident_energy_layout.addLayout(self.high_energy_layout)
        spacerItem2 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.incident_energy_layout.addItem(spacerItem2)
        self.arc_flash_grid.addLayout(self.incident_energy_layout, 1, 2, 1, 1)
        self.verticalLayout_11.addLayout(self.arc_flash_grid)
        self.main_layout.addWidget(self.arc_flash_layout)
        self.etap_dir_group = QtWidgets.QGroupBox(self.centralwidget)
        self.etap_dir_group.setEnabled(True)
        self.etap_dir_group.setObjectName("etap_dir_group")
        self.verticalLayout_8 = QtWidgets.QVBoxLayout(self.etap_dir_group)
        self.verticalLayout_8.setObjectName("verticalLayout_8")
        self.etap_dir = QtWidgets.QLineEdit(self.etap_dir_group)
        self.etap_dir.setMinimumSize(QtCore.QSize(0, 24))
        self.etap_dir.setReadOnly(False)
        self.etap_dir.setClearButtonEnabled(True)
        self.etap_dir.setObjectName("etap_dir")
        self.verticalLayout_8.addWidget(self.etap_dir)
        self.browse_btn_layout = QtWidgets.QHBoxLayout()
        self.browse_btn_layout.setObjectName("browse_btn_layout")
        self.browse_btn = QtWidgets.QPushButton(self.etap_dir_group)
        self.browse_btn.setObjectName("browse_btn")
        self.browse_btn_layout.addWidget(self.browse_btn)
        spacerItem3 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.browse_btn_layout.addItem(spacerItem3)
        self.verticalLayout_8.addLayout(self.browse_btn_layout)
        self.main_layout.addWidget(self.etap_dir_group)
        self.output_dir_group = QtWidgets.QGroupBox(self.centralwidget)
        self.output_dir_group.setObjectName("output_dir_group")
        self.verticalLayout_9 = QtWidgets.QVBoxLayout(self.output_dir_group)
        self.verticalLayout_9.setObjectName("verticalLayout_9")
        self.output_dir = QtWidgets.QLineEdit(self.output_dir_group)
        self.output_dir.setMinimumSize(QtCore.QSize(0, 24))
        self.output_dir.setClearButtonEnabled(True)
        self.output_dir.setObjectName("output_dir")
        self.verticalLayout_9.addWidget(self.output_dir)
        self.browse_btn_layout_2 = QtWidgets.QHBoxLayout()
        self.browse_btn_layout_2.setObjectName("browse_btn_layout_2")
        self.browse_btn_2 = QtWidgets.QPushButton(self.output_dir_group)
        self.browse_btn_2.setObjectName("browse_btn_2")
        self.browse_btn_layout_2.addWidget(self.browse_btn_2)
        self.etap_dir_checkbox = QtWidgets.QCheckBox(self.output_dir_group)
        self.etap_dir_checkbox.setObjectName("etap_dir_checkbox")
        self.browse_btn_layout_2.addWidget(self.etap_dir_checkbox)
        spacerItem4 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.browse_btn_layout_2.addItem(spacerItem4)
        self.verticalLayout_9.addLayout(self.browse_btn_layout_2)
        self.main_layout.addWidget(self.output_dir_group)
        self.exclude_group = QtWidgets.QGroupBox(self.centralwidget)
        self.exclude_group.setObjectName("exclude_group")
        self.verticalLayout_4 = QtWidgets.QVBoxLayout(self.exclude_group)
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.exclude_start_layout = QtWidgets.QVBoxLayout()
        self.exclude_start_layout.setObjectName("exclude_start_layout")
        self.exclude_start_label = QtWidgets.QLabel(self.exclude_group)
        self.exclude_start_label.setObjectName("exclude_start_label")
        self.exclude_start_layout.addWidget(self.exclude_start_label)
        self.verticalLayout_4.addLayout(self.exclude_start_layout)
        self.exclude_start_input = QtWidgets.QLineEdit(self.exclude_group)
        self.exclude_start_input.setMinimumSize(QtCore.QSize(0, 24))
        self.exclude_start_input.setObjectName("exclude_start_input")
        self.verticalLayout_4.addWidget(self.exclude_start_input)
        self.exclude_contain_layout = QtWidgets.QVBoxLayout()
        self.exclude_contain_layout.setObjectName("exclude_contain_layout")
        self.exclude_contain_label = QtWidgets.QLabel(self.exclude_group)
        self.exclude_contain_label.setObjectName("exclude_contain_label")
        self.exclude_contain_layout.addWidget(self.exclude_contain_label)
        self.verticalLayout_4.addLayout(self.exclude_contain_layout)
        self.exclude_contain_input = QtWidgets.QLineEdit(self.exclude_group)
        self.exclude_contain_input.setMinimumSize(QtCore.QSize(0, 24))
        self.exclude_contain_input.setObjectName("exclude_contain_input")
        self.verticalLayout_4.addWidget(self.exclude_contain_input)
        self.main_layout.addWidget(self.exclude_group)
        self.cntrl_layout = QtWidgets.QVBoxLayout()
        self.cntrl_layout.setObjectName("cntrl_layout")
        self.datahub_note_layout = QtWidgets.QVBoxLayout()
        self.datahub_note_layout.setContentsMargins(2, 0, -1, 0)
        self.datahub_note_layout.setObjectName("datahub_note_layout")
        self.datahub_note = QtWidgets.QLabel(self.centralwidget)
        self.datahub_note.setEnabled(True)
        self.datahub_note.setLineWidth(1)
        self.datahub_note.setAlignment(QtCore.Qt.AlignLeading|QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        self.datahub_note.setObjectName("datahub_note")
        self.datahub_note_layout.addWidget(self.datahub_note)
        self.cntrl_layout.addLayout(self.datahub_note_layout)
        self.cntrl_btns_layout = QtWidgets.QHBoxLayout()
        self.cntrl_btns_layout.setContentsMargins(0, 6, -1, -1)
        self.cntrl_btns_layout.setObjectName("cntrl_btns_layout")
        self.clear_all_btn = QtWidgets.QPushButton(self.centralwidget)
        self.clear_all_btn.setObjectName("clear_all_btn")
        self.cntrl_btns_layout.addWidget(self.clear_all_btn)
        spacerItem5 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.cntrl_btns_layout.addItem(spacerItem5)
        self.generate_btn = QtWidgets.QPushButton(self.centralwidget)
        self.generate_btn.setObjectName("generate_btn")
        self.cntrl_btns_layout.addWidget(self.generate_btn)
        self.close_btn = QtWidgets.QPushButton(self.centralwidget)
        self.close_btn.setObjectName("close_btn")
        self.cntrl_btns_layout.addWidget(self.close_btn)
        self.cntrl_layout.addLayout(self.cntrl_btns_layout)
        self.main_layout.addLayout(self.cntrl_layout)
        self.verticalLayout_2.addLayout(self.main_layout)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 448, 21))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuEdit = QtWidgets.QMenu(self.menubar)
        self.menuEdit.setObjectName("menuEdit")
        self.menuHelp = QtWidgets.QMenu(self.menubar)
        self.menuHelp.setObjectName("menuHelp")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuEdit.menuAction())
        self.menubar.addAction(self.menuHelp.menuAction())

        self.retranslateUi(MainWindow)
        self.close_btn.clicked.connect(MainWindow.close) # type: ignore
        self.create_scenarios_checkbox.toggled['bool'].connect(self.etap_dir_group.setDisabled) # type: ignore
        self.create_scenarios_checkbox.toggled['bool'].connect(self.etap_dir.clear) # type: ignore
        self.create_scenarios_checkbox.toggled['bool'].connect(self.run_scenarios_checkbox.setEnabled) # type: ignore
        self.create_scenarios_checkbox.toggled['bool'].connect(self.run_scenarios_checkbox.setChecked) # type: ignore
        self.device_duty_checkbox.toggled['bool'].connect(self.device_duty_group.setEnabled) # type: ignore
        self.etap_dir_checkbox.toggled['bool'].connect(self.output_dir.setDisabled) # type: ignore
        self.etap_dir_checkbox.toggled['bool'].connect(self.browse_btn_2.setDisabled) # type: ignore
        self.arc_flash_checkbox.toggled['bool'].connect(self.arc_flash_layout.setEnabled) # type: ignore
        self.device_duty_checkbox.toggled['bool'].connect(self.mark_assumed_checkbox.setChecked) # type: ignore
        self.device_duty_checkbox.toggled['bool'].connect(self.series_rating_checkbox.setChecked) # type: ignore
        self.device_duty_checkbox.toggled['bool'].connect(self.sw_checkbox.setChecked) # type: ignore
        self.device_duty_checkbox.toggled['bool'].connect(self.swgr_checkbox.setChecked) # type: ignore
        self.create_scenarios_checkbox.toggled['bool'].connect(self.nm_switching_radio.setEnabled) # type: ignore
        self.create_scenarios_checkbox.toggled['bool'].connect(self.nm_revisions_radio.setEnabled) # type: ignore
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "PSS Table Generator"))
        self.title.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" font-size:11pt;\">PSS Table Generator</span></p></body></html>"))
        self.port_group.setTitle(_translate("MainWindow", "Datahub Port"))
        self.port.setText(_translate("MainWindow", "65358"))
        self.options_group.setTitle(_translate("MainWindow", "Options"))
        self.create_scenarios_checkbox.setText(_translate("MainWindow", "Create Scenarios *"))
        self.run_scenarios_checkbox.setText(_translate("MainWindow", "Run Scenarios"))
        self.report_group.setTitle(_translate("MainWindow", "Report"))
        self.arc_flash_checkbox.setText(_translate("MainWindow", "Arc Flash"))
        self.device_duty_checkbox.setText(_translate("MainWindow", "Device Duty"))
        self.device_duty_group.setTitle(_translate("MainWindow", "Device Duty"))
        self.sw_checkbox.setText(_translate("MainWindow", "Add Switch Asymm."))
        self.swgr_checkbox.setText(_translate("MainWindow", "Add Switchgear Symm."))
        self.mark_assumed_checkbox.setText(_translate("MainWindow", "Mark Assumed Equipment *"))
        self.series_rating_checkbox.setText(_translate("MainWindow", "Add Series Ratings *"))
        self.arc_flash_layout.setTitle(_translate("MainWindow", "Arc Flash"))
        self.nm_config_label.setText(_translate("MainWindow", "No Motor Configs:    "))
        self.nm_switching_radio.setText(_translate("MainWindow", "Use Switching"))
        self.nm_revisions_radio.setText(_translate("MainWindow", "Use Revisions"))
        self.incident_energy_label.setText(_translate("MainWindow", "Incident Energy:"))
        self.low_energy_label.setText(_translate("MainWindow", "High"))
        self.high_energy_label.setText(_translate("MainWindow", "Critical"))
        self.etap_dir_group.setTitle(_translate("MainWindow", "ETAP Directory"))
        self.etap_dir.setPlaceholderText(_translate("MainWindow", "Path to Project ETAP Folder"))
        self.browse_btn.setText(_translate("MainWindow", "Browse"))
        self.output_dir_group.setTitle(_translate("MainWindow", "Output Directory"))
        self.output_dir.setPlaceholderText(_translate("MainWindow", "Path to Output Folder"))
        self.browse_btn_2.setText(_translate("MainWindow", "Browse"))
        self.etap_dir_checkbox.setText(_translate("MainWindow", "Use ETAP Directory"))
        self.exclude_group.setTitle(_translate("MainWindow", "Exclude"))
        self.exclude_start_label.setText(_translate("MainWindow", "Element ID Starting With:"))
        self.exclude_start_input.setText(_translate("MainWindow", "BUS;"))
        self.exclude_contain_label.setText(_translate("MainWindow", "Element ID Containing:"))
        self.exclude_contain_input.setText(_translate("MainWindow", "BCH;FORTIS;"))
        self.datahub_note.setText(_translate("MainWindow", "<html><head/><body><p><span style=\" color:#ff0000;\">* Note: Selected inputs require ETAP Datahub running.</span></p></body></html>"))
        self.clear_all_btn.setText(_translate("MainWindow", "Clear All"))
        self.generate_btn.setText(_translate("MainWindow", "Generate"))
        self.close_btn.setText(_translate("MainWindow", "Close"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuEdit.setTitle(_translate("MainWindow", "Edit"))
        self.menuHelp.setTitle(_translate("MainWindow", "Help"))
