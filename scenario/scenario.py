import etap.api
import json
from pathlib import Path
import xml.etree.ElementTree as ET
from consts.consts import SYSTEM, PRESENTATION


class Scenario:
    """
    The Scenario class handles the creation, management, and execution of scenarios
    within the ETAP environment using an XML configuration file.
    """

    def __init__(self, study_mode: str, port: int = 65358):
        """
        Initializes a Scenario instance with the specified study mode and ETAP connection port.

        :param str study_mode: The study mode to be used for the scenario (e.g., "ANSI", "IEC").
        :param int port: The port number for connecting to the ETAP Datahub (default is 65358).
        """
        self.etap = etap.api.connect(f'http://localhost:{port}')
        self.scenario_ids = []
        self.scenario_xml = ET.parse(self.get_scenario_xml_path())
        self.scenario_root = self.scenario_xml.getroot()
        self.study_mode = study_mode

    def run_scenarios(self):
        """
        Runs all scenarios listed in self.scenario_ids by invoking ETAP's scenario run method.
        """
        for scenario_id in self.scenario_ids:
            self.etap.scenario.run(scenario_id)

    def get_project_path(self) -> Path:
        """
        Retrieves the full path of the current ETAP project.

        :return: The full path to the current ETAP project as a Path object.
        :rtype: Path
        """
        return Path(json.loads(self.etap.application.projectfile())['FullPath'])

    def get_project_dir(self) -> Path:
        """
        Retrieves the directory of the current ETAP project.

        :return: The directory of the current ETAP project as a Path object.
        :rtype: Path
        """
        return self.get_project_path().parent

    def get_scenario_xml_path(self) -> Path:
        """
        Constructs the path to the scenarios XML file based on the current project path.

        :return: The path to the scenarios XML file as a Path object.
        :rtype: Path
        """
        return self.get_project_dir() / (self.get_project_path().stem + '.scenarios.xml')

    def write_scenario_xml(self):
        """
        Writes the current state of the scenario XML tree to the scenarios XML file.
        """
        self.scenario_xml.write(self.get_scenario_xml_path())

    def create_scenario(self, scenario_id: str, switching_config: str, study_case: str,
                        rev_config: str, output: str):
        """
        Creates a new scenario in the XML file if it does not already exist.

        :param str scenario_id: The unique identifier for the scenario.
        :param str switching_config: The switching configuration for the scenario.
        :param str study_case: The study case configuration for the scenario.
        :param str rev_config: The revision configuration for the scenario.
        :param str output: The output file name for the scenario results.
        """
        # Check if the scenario already exists
        if self.scenario_root.findall(f'./Scenario[@ID="{scenario_id}"]'):
            return

        # Create a new scenario element with the specified attributes
        scenario_element = ET.Element('Scenario')
        scenario_element.set('ID', scenario_id)
        scenario_element.set('Executable', 'Yes')
        scenario_element.set('ForceSave', 'Yes')
        scenario_element.set('background', 'No')
        scenario_element.set('ToolTip', r'\n\n')
        scenario_element.set('System', SYSTEM)
        scenario_element.set('Presentation', PRESENTATION)
        scenario_element.set('Mode', f'STUDY_SHORTCIRCUIT {self.study_mode}')
        scenario_element.set('Config', switching_config)
        scenario_element.set('StudyCase', study_case)
        scenario_element.set('Revision', rev_config)
        scenario_element.set('Output', output)
        scenario_element.set('Sequence', "")
        scenario_element.set('ActionTool', f'STUDY_SHORTCIRCUIT {self.study_mode}')
        scenario_element.set('Comments', "")
        scenario_element.set('Compare', "False")
        scenario_element.set('NewFilePath', rf'.\{output}.AAFS')
        scenario_element.set('Compare_benchmark', "")
        scenario_element.set('InstructionFilePath', "")
        scenario_element.set('UseETAPDefaultLibrary', "False")
        scenario_element.set('Compare_deviationReportFile', "")
        scenario_element.set('Compare_globalSummaryFile', "")
        scenario_element.set('Compare_skipRecordsThatPass', "True")
        scenario_element.set('Compare_percentDev', "0.1")
        scenario_element.set('Compare_skipProjInfo', "True")
        scenario_element.set('Compare_remarks', "")
        scenario_element.set('Compare_commandLine', "")
        scenario_element.set('Compare_skipDates', "")
        scenario_element.set('Compare_autoOpen', "")
        scenario_element.set('IniSettings', "")
        scenario_element.set('GetOnlineData', "False")
        scenario_element.set('WhatIfCommands', "")
        scenario_element.set('IsComparePlot', "")
        scenario_element.set('PlotCompareOutput', "")
        scenario_element.set('MaxPlotDiff', "")
        scenario_element.set('TotalPlotDiff', "")
        scenario_element.set('ConfigDBID', "")
        scenario_element.set('PresentationDBID', "")
        scenario_element.set('RevisionDBID', "")
        scenario_element.set('StudyCaseDBID', "")
        scenario_element.set('UseRealTime', "True")
        scenario_element.set('GetArchiveForOTS', "False")
        scenario_element.set('UseArchived', "False")
        scenario_element.set('UseRTConfig', "False")
        scenario_element.set('UseArchivedConfig', "False")

        # Append the new scenario element to the root of the XML tree
        self.scenario_root.append(scenario_element)

