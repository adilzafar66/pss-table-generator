# ***********************
#
# Copyright (c) 2021-2023, Operation Technology, Inc.
#
# THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO OPERATION TECHNOLOGY, INC.
# ANY USE OF THIS PROGRAM IS SUBJECT TO THE PROGRAM SOFTWARE LICENSE AGREEMENT,
# EXCEPT THAT THE USER MAY MODIFY THE PROGRAM FOR ITS OWN USE.
# HOWEVER, THE PROGRAM MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS
# WITHOUT THE PRIOR WRITTEN CONSENT OF OPERATION TECHNOLOGY, INC.
#
# ***********************


from .other import datahub as datahub
from .other import api_constants as apiconstants
import urllib.request
import urllib.parse


class Scenario:

    def __init__(self, baseAddress: str, sentToken: str, projectName=None):
        self._baseAddress = baseAddress
        self._token = sentToken
        self._projectName = projectName

    def getxmlfilepath(self) -> str:
        """Returns the location of an XML file containing scenario information.
        If on the same machine, the file can be opened directly.  If on another machine
        the file must be downloaded using datahub.get_file()"""
        return datahub.get(f'{self._baseAddress}{apiconstants.scenario_getxmlfilepath}', token=self._token)

    def getxml(self) -> str:
        """Returns XML describing all scenarios defined for the current project."""
        return datahub.get(f'{self._baseAddress}{apiconstants.scenario_getxml}', token=self._token)

    def run(self, id: str, getOnlineData=None, whatIfCommands=None) -> str:
        """Runs the specified scenario.  Nothing is returned."""

        httpLocation = f'{self._baseAddress}{apiconstants.scenario_run}?'
        params = "id={0}&getOnlineData={1}".format(id, getOnlineData)
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        return datahub.post(address, whatIfCommands, token=self._token)

    def createscenario(self, scenarioID: str, system: str, presentation: str, revisionName: str, configName: str, studyMode: str, studyType: str, studyCase: str, outputReport: str, whatIfCommands=None) -> str:
        """Creates the specified scenario.  Returns boolean True or False."""
        try:
            # datahub.post(f'{self._baseAddress}{apiconstants.scenario_createscenario}/{scenarioID}/{system}/{presentation}/{revisionName}/{configName}/{studyMode}/{studyType}/{studyCase}/{outputReport}', whatIfCommands, token=self._token)
            httpLocation = f'{self._baseAddress}{apiconstants.scenario_createscenario}?'
            param_list = [scenarioID, system, presentation, revisionName, configName, studyMode, studyType, studyCase, outputReport]
        
            # URL encode all the params
            for elem in param_list:
                param_list[param_list.index(elem)] = urllib.parse.quote(elem)
            
            params = f"scenarioID={param_list[0]}&system={param_list[1]}&presentation={param_list[2]}&revisionName={param_list[3]}&configName={param_list[4]}&studyMode={param_list[5]}&studyType={param_list[6]}&studyCase={param_list[7]}&outputReport={param_list[8]}"
            address = httpLocation + params

            datahub.post(address, whatIfCommands, token=self._token)

            return True
        except:
            return False

    def checkbeforeoperate(self, id: str, getOnlineData=None, whatIfCommands=None) -> str:
        """Runs the specified scenario.  Check alerts and de-engerized elements affected by this scenario.."""

        httpLocation = f'{self._baseAddress}{apiconstants.scenario_checkbeforeoperate}?'
        params = "id={0}&getOnlineData={1}".format(id, getOnlineData)
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        return datahub.post(address, whatIfCommands, token=self._token)

    class WhatifCommands:

        def __init__(self, baseAddress: str, sentToken: str, projectName=None):
            self._baseAddress = baseAddress
            self._token = sentToken
            self._projectName = projectName

        def buskvar(self, id: str, whatIfCommands) -> str:
            """Runs the specified whatifcommand.  Alerts are returned."""

            httpLocation = f'{self._baseAddress}{apiconstants.scenario_whatif_buskvar}?'
            params = "id={0}".format(id)
            address = httpLocation + urllib.parse.quote(params, safe='&=')

            return datahub.post(address, whatIfCommands, token=self._token)

        def buskw(self, id: str, whatIfCommands) -> str:
            """Runs the specified whatifcommand.  Alerts are returned."""

            httpLocation = f'{self._baseAddress}{apiconstants.scenario_whatif_buskw}?'
            params = "id={0}".format(id)
            address = httpLocation + urllib.parse.quote(params, safe='&=')

            return datahub.post(address, whatIfCommands, token=self._token)

        def close(self, id: str, whatIfCommands) -> str:
            """Runs the specified whatifcommand.  Alerts are returned."""

            httpLocation = f'{self._baseAddress}{apiconstants.scenario_whatif_close}?'
            params = "id={0}".format(id)
            address = httpLocation + urllib.parse.quote(params, safe='&=')

            return datahub.post(address, whatIfCommands, token=self._token)

        def open(self, id: str, whatIfCommands) -> str:
            """Runs the specified whatifcommand.  Alerts are returned."""

            httpLocation = f'{self._baseAddress}{apiconstants.scenario_whatif_open}?'
            params = "id={0}".format(id)
            address = httpLocation + urllib.parse.quote(params, safe='&=')

            return datahub.post(address, whatIfCommands, token=self._token)
