#***********************
#
# Copyright (c) 2021-2023, Operation Technology, Inc.
#
# THIS PROGRAM IS CONFIDENTIAL AND PROPRIETARY TO OPERATION TECHNOLOGY, INC. 
# ANY USE OF THIS PROGRAM IS SUBJECT TO THE PROGRAM SOFTWARE LICENSE AGREEMENT, 
# EXCEPT THAT THE USER MAY MODIFY THE PROGRAM FOR ITS OWN USE. 
# HOWEVER, THE PROGRAM MAY NOT BE REPRODUCED, PUBLISHED, OR DISCLOSED TO OTHERS 
# WITHOUT THE PRIOR WRITTEN CONSENT OF OPERATION TECHNOLOGY, INC.
#
#***********************




import urllib.request
import re
import json
import xml.etree.ElementTree as ET
import urllib.parse
import base64
import os
import sys
import re
from .other import datahub as datahub
from .other import api_constants as apiconstants
import tempfile
import requests
import traceback

class Studies:
    """Contains the possible ETAP studies that one can run from Python."""

    def __init__(self, baseAddress: str, isForRemoteEtap: bool, sentToken:str, projectName=None):
        self._baseAddress = baseAddress
        self._etapClient = None
        self._isForRemoteEtap = isForRemoteEtap
        self._token = sentToken
        self._projectName = projectName

    # def __init__(self, etapClient):
    #    """Constructor"""
    #    self._etapClient = etapClient

    def runULF(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputReport: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a ULF study in ETAP.  Returns the report name.  The results are available on the one-line diagram and in the output database."""

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension
        
        # address
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runulf}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData, onlineConfigOnly)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        try: 
            if (revisionName and configName and studyCase and presentation and outputReport):
                try:
                    result = datahub.post(address, whatIfCommands, token=self._token)
                    return result

                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputReport):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
#                    raise ValueError('RevisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
#                    raise ValueError('ConfigName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputReport:
                    print('outputReport parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()


        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------


    def runLF(self, revisionName: str, configName: str, studyCase: str, presentation: str, 
              outputReport: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a LF study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation to run the study in
        5. 'outputReport'              : The output file for the study
        6. 'getOnlineData              : Whether or not to Get Online Data for the study
        7. 'onlineConfigOnly'          : Whether or not to Get online configurations for the study
        8. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')
        """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        # NOTE: parameter 6 above is a dictionary which is submitted as the body
        # of the POST operation (i.e., not part of the URI)
        # address = "http://{0}:{1}/pythonservice2/{2}/runlf2/{3}/{4}/{5}/{6}/{7}/{8}/{9}".format(self._ipAddressOrComputerName, self._portNumber, self._projectNameNoExtension, revisionName, configName, studyCase, outputReport, getOnlineData, whatIfMotorOperatingLoad, whatIfGenGeneration)
        # httpLocation = "http://{0}:{1}/pythonservice2/{2}/runlf2/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runlf}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData, onlineConfigOnly)
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        try: 
            if (revisionName and configName and studyCase and presentation and outputReport):
                try:
                    result = datahub.post(address, whatIfCommands, token=self._token)
                    return result
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputReport):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
#                    raise ValueError('RevisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
#                    raise ValueError('ConfigName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputReport:
                    print('outputReport parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------
        

    def runTSsync(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a TS study in ETAP.  Returns the report name.  The results are available on the
           one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'presentation'              : The presentation to run the study in
            5. 'outputFile                 : The output file name for the study results
            6. 'getOnlineData'             : Whether or not to get online data for the study
            7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
            8. 'runAsync'                  : Whether or not to run asynchronously
            9. 'whatIfCommands'            : Dictionary of what-if commands
            """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        #httpLocation = "http://{0}:{1}/pythonservice2/{2}/runts/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runtssync}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputFile, getOnlineData, onlineConfigOnly)
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')
        

        try: 
            if (revisionName and configName and studyCase and presentation and outputFile):
                try:
                    result = datahub.post(address, whatIfCommands, token=self._token)
                    return result
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputFile):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
#                    raise ValueError('RevisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
#                    raise ValueError('ConfigName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputFile:
                    print('outputFile parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------
    def runTSasync(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a TS study in ETAP.  Returns the report name.  The results are available on the
           one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'presentation'              : The presentation to run the study in
            5. 'outputFile                 : The output file name for the study results
            6. 'getOnlineData'             : Whether or not to get online data for the study
            7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
            8. 'runAsync'                  : Whether or not to run asynchronously
            9. 'whatIfCommands'            : Dictionary of what-if commands
            """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        #httpLocation = "http://{0}:{1}/pythonservice2/{2}/runts/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runtsasync}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputFile, getOnlineData, onlineConfigOnly)
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        req = urllib.request.Request(address)
        req.add_header('Content-Type', 'application/json; charset=utf-8')

        try: 
            if (revisionName and configName and studyCase and presentation and outputFile):
                try:
                    result = datahub.post(address, whatIfCommands, token=self._token)
                    return result
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputFile):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
#                    raise ValueError('RevisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
#                    raise ValueError('ConfigName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputFile:
                    print('outputFile parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------
        
    

    # def runTSAsync(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, getOnlineData: bool, whatIf1: str, whatIf2: str, whatIf3: str) -> str:
    #     """Runs a TS study in ETAP.  *Does not return* the report name like RunTS because this method runs asynchronously.
    #         1. 'revisionName'              : The revision for the study
    #         2. 'configName'                : The config for the study
    #         3. 'studyCase'                 : The study case for the study
    #         4. 'presentation'              : The presentation to run the study in
    #         5. 'outputFile                 : The output file name for the study results
    #         6. 'getOnlineData'             : Whether or not to get online data for the study
    #         7. 'whatIf1'                   : What-ifs that should be use as overrides for the study
    #         8. 'whatIf2'                   : N/A
    #         9. 'whatIf3'                   : N/A"""

    #     # get connection information
    #     ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
    #     portNumber = self._etapClient._portNumber
    #     projectNameNoExtension = self._etapClient._projectNameNoExtension

    #     httpLocation = "http://{0}:{1}/pythonservice2/{2}/runts/".format(
    #         ipAddressOrComputerName, portNumber, projectNameNoExtension)
    #     params = "{0}/{1}/{2}/{3}/{4}/{5}/{6}/{7}/{8}".format(
    #         revisionName, configName, studyCase, presentation, outputFile, getOnlineData, whatIf2, whatIf3, True)
    #     # url encode params before sending
    #     address = httpLocation + urllib.parse.quote(params)

    #     req = urllib.request.Request(address)
    #     req.add_header('Content-Type', 'application/json; charset=utf-8')

    #     # Convert 'whatIfPdStatusDict' to a JSON string which goes in the body
    #     # of the POST operation
    #     sendThisString = json.dumps(whatIf1,
    #                                 indent=4, sort_keys=False,
    #                                 separators=(',', ':'), ensure_ascii=False)

    #     # Convert POST body to bytes
    #     jsonBytes = sendThisString.encode('utf-8')

    #     req.add_header('Content-Length', len(jsonBytes))
    #     # ------------------------------------------
    #     try:
    #         # OK for runTsAsync to run remotely.  This call does not return a result  locally nor remotely.
    #         # if self._etapClient._isForRemoteEtap == True:
    #         #    raise ValueError
    #         # h = urllib.request.urlopen(address)
    #         h = urllib.request.urlopen(req, jsonBytes)
    #         xmlString = "".join(map(chr, h.read()))
    #         xmlString = self.__removeUnwantedChars(xmlString)
    #         xmlString = self.__decodeUnicodeEncodedString(xmlString)
    #         xmlString = self.__removeXmlVersionTag(xmlString)
    #         return xmlString
    #     except KeyboardInterrupt:
    #         print('Operation cancelled.')
    #     # except ValueError:
    #     #    print("Error: retrieving simulation results from a remote machine is only supported by the ‘runTs’  function.  To work around this limitation, run Python and ETAP on the same machine.")
    #     except:
    #         print('Error: could not run TS (python-side exception).')

    def runMS(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, numMotorStart: str, studyType: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a MS study in ETAP.  Returns the report name.  The results are available on the
           one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'presentation'              : The presentation to run the study in
            5. 'outputFile                 : The output file name for the study results
            6. 'getOnlineData'             : Whether or not to get online data for the study
            7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
            8. 'numMotorStart'             : 0 = start all, 1 = start 1 at a time, 2 = start 2 at a time, 3 = start 3 at a time, ..., n = start n at a time
            9. 'studyType'                 : "Dynamic" or "Static"
            10. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')"""

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        #httpLocation = "http://{0}:{1}/pythonservice2/{2}/runms/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runms}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}&numMotorStart={7}&studyType={8}".format(
            revisionName, configName, studyCase, presentation, outputFile, getOnlineData, onlineConfigOnly, numMotorStart, studyType)

        address = httpLocation + urllib.parse.quote(params, safe='&=')

        try: 
            if (revisionName and configName and studyCase and presentation and outputFile and numMotorStart and studyType):
                try:
                    result = datahub.post(address, whatIfCommands, token=self._token)
                    return result
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputFile and numMotorStart and studyType):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
#                    raise ValueError('RevisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
#                    raise ValueError('ConfigName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputFile:
                    print('outputFile parameter is required. It should not be empty')
                if not studyType:
                    print('studyType parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()


    def runTDLF(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a TDLF study in ETAP.  Returns the report name.  The results are available on the
           one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'presentation'              : The presentation to run the study in
            5. 'outputFile                 : The output file name for the study results
            6. 'getOnlineData'             : Whether or not to get online data for the study
            7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
            8. 'whatIfCommands'                    : Dictionary of what-if commands
        """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        #httpLocation = "http://{0}:{1}/pythonservice2/{2}/runtdlf/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runtdlf}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputFile, getOnlineData, onlineConfigOnly)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')
        
        result = datahub.post(address, whatIfCommands, token=self._token)

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------

        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine.')
            # TODO: -Neetin Choudary 2/16/2022
            # ipAddress is not passed and need to be tested for remote python api. This effects multiple studies.
            # Current scenario assumes base address as localhost only
            # print('Unable to download file from the remote machine ({}).'.format(
            #     ipAddressOrComputerName))
        except:
            print('Error: could not run TDLF (python-side exception).')

    
    
    def runTDLFasync(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a time-domain load flow study asynchronously. The output report location is returned in the response body.
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'presentation'              : The presentation to run the study in
            5. 'outputFile                 : The output file name for the study results
            6. 'getOnlineData'             : Whether or not to get online data for the study
            7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
            8. 'whatIfCommands'            : Dictionary of what-if commands
        """

        httpLocation = f'{self._baseAddress}{apiconstants.studies_runtdlfasync}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputFile, getOnlineData, onlineConfigOnly)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        result = datahub.post(address, whatIfCommands, token=self._token)

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------

        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine.')
            # TODO: -Neetin Choudary 2/16/2022
            # ipAddress is not passed and need to be tested for remote python api. This effects multiple studies.
            # Current scenario assumes base address as localhost only
            # print('Unable to download file from the remote machine ({}).'.format(
            #     ipAddressOrComputerName))
        except:
            print('Error: could not run TDLF (python-side exception).')


    def runVS(self, revisionName: str, configName: str, studyCase: str, presentation: str,
              outputReport: str, studyType: str, getOnlineData=None, whatIfCommands=None) -> str:
        """Runs a VS study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation for the study
        5. 'outputReport'              : The output file for the study
        6. 'getOnlineData              : Whether or not to Get Online Data for the study
        7. 'studyType'                 : Which VS analysis to run
        8. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')
        """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        # NOTE: parameter 6 above is a dictionary which is submitted as the body
        # of the POST operation (i.e., not part of the URI)
        # address = "http://{0}:{1}/pythonservice2/{2}/runlf2/{3}/{4}/{5}/{6}/{7}/{8}/{9}".format(self._ipAddressOrComputerName, self._portNumber, self._projectNameNoExtension, revisionName, configName, studyCase, outputReport, getOnlineData, whatIfMotorOperatingLoad, whatIfGenGeneration)
        # httpLocation = "http://{0}:{1}/pythonservice2/{2}/runlf2/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runvs}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&studyType={6}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData, studyType)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')
        
        result = datahub.post(address, whatIfCommands, token=self._token)
        
        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------
        
        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine.')
            # TODO: -Neetin Choudary 2/16/2022
            # ipAddress is not passed and need to be tested for remote python api. This effects multiple studies.
            # Current scenario assumes base address as localhost only
            # print('Unable to download file from the remote machine ({}).'.format(
            #     ipAddressOrComputerName))
        except:
            print('Error: could not run VS (python-side exception).')


    def runHA(self, revisionName: str, configName: str, studyCase: str, presentation: str,
              outputReport: str, studyType: str, getOnlineData=None, whatIfCommands=None) -> str:
        """Runs a HA study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation for the study
        5. 'outputReport'              : The output file for the study
        6. 'getOnlineData              : Whether or not to Get Online Data for the study
        7. 'studyType'                 : Which HA analysis to run
        8. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')
        """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        # NOTE: parameter 6 above is a dictionary which is submitted as the body
        # of the POST operation (i.e., not part of the URI)
        # address = "http://{0}:{1}/pythonservice2/{2}/runlf2/{3}/{4}/{5}/{6}/{7}/{8}/{9}".format(self._ipAddressOrComputerName, self._portNumber, self._projectNameNoExtension, revisionName, configName, studyCase, outputReport, getOnlineData, whatIfMotorOperatingLoad, whatIfGenGeneration)
        # httpLocation = "http://{0}:{1}/pythonservice2/{2}/runlf2/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runha}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&studyType={6}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData, studyType)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')     

        result = datahub.post(address, whatIfCommands, token=self._token)
        
        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------
        
        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine.')
            # TODO: -Neetin Choudary 2/16/2022
            # ipAddress is not passed and need to be tested for remote python api. This effects multiple studies.
            # Current scenario assumes base address as localhost only
            # print('Unable to download file from the remote machine ({}).'.format(
            #     ipAddressOrComputerName))
        except:
            print('Error: could not run HA (python-side exception).')

    def runSC(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputFile: str, studyType: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a TDLF study in ETAP.  Returns the report name.  The results are available on the
           one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'presentation'              : The presentation to run the study in
            5. 'outputFile                 : The output file name for the study results
            6. 'getOnlineData'             : Whether or not to get online data for the study
            7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
            8. 'studyType'                 : Which SC analysis to run
            9. 'whatIfCommands'            : Dictionary of what-if commands
        """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        #httpLocation = "http://{0}:{1}/pythonservice2/{2}/runtdlf/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runsc}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}&studyType={7}".format(
            revisionName, configName, studyCase, presentation, outputFile, getOnlineData, onlineConfigOnly, studyType)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')
        
        result = datahub.post(address, whatIfCommands, token=self._token)

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------

        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine.')
            # TODO: -Neetin Choudary 2/16/2022
            # ipAddress is not passed and need to be tested for remote python api. This effects multiple studies.
            # Current scenario assumes base address as localhost only
            # print('Unable to download file from the remote machine ({}).'.format(
            #     ipAddressOrComputerName))
        except:
            print('Error: could not run TDLF (python-side exception).')


    def runStarSQOP(self, revisionName: str, configName: str, studyCase: str, presentation: str,
              outputReport: str, busID: str, faultType: str, whatIfCommands=None) -> str:
        """Runs a HA study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation for the study
        5. 'outputReport'              : The output file for the study
        6. 'busID'                     : Energized Bus ID
        7. 'faultType'                 : Fault type (e.g., "3-Phase", "Line-to-Ground", "Line-to-Line", "Line-to-Line-to-Ground")
        """

        # get connection information
        #ipAddressOrComputerName = self._etapClient._ipAddressOrComputerName
        #portNumber = self._etapClient._portNumber
        #projectNameNoExtension = self._etapClient._projectNameNoExtension

        # NOTE: parameter 6 above is a dictionary which is submitted as the body
        # of the POST operation (i.e., not part of the URI)
        # address = "http://{0}:{1}/pythonservice2/{2}/runlf2/{3}/{4}/{5}/{6}/{7}/{8}/{9}".format(self._ipAddressOrComputerName, self._portNumber, self._projectNameNoExtension, revisionName, configName, studyCase, outputReport, getOnlineData, whatIfMotorOperatingLoad, whatIfGenGeneration)
        # httpLocation = "http://{0}:{1}/pythonservice2/{2}/runlf2/".format(
        #    ipAddressOrComputerName, portNumber, projectNameNoExtension)
        # httpLocation = f'{self._baseAddress}{apiconstants.studies_runstarsqop}/'
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runstarsqop}?'

        param_list = [revisionName, configName, studyCase, presentation, outputReport, busID, faultType]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"revisionName={param_list[0]}&configName={param_list[1]}&studyCase={param_list[2]}&presentation={param_list[3]}&outputReport={param_list[4]}&busID={param_list[5]}&faultType={param_list[6]}"
        address = httpLocation + params

        # params = "{0}/{1}/{2}/{3}/{4}/{5}/{6}".format(revisionName, configName, studyCase, presentation,
        #                                       outputReport, busID, faultType)

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------

        try: 
            if (revisionName and configName and studyCase and presentation and outputReport and busID and faultType):
                try:
                    result = datahub.post(address, whatIfCommands, token=self._token)
                    import json
                    P_Dict = json.loads(result)  # Python Dict
                    #print(P_Dict)
                    P_Value = P_Dict["SQOPEventFile"]
                    return P_Value
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
#                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
                    return exception.response.status_code     
#                    return "ETAP fails to run Star SQOP"
            if not (revisionName and configName and studyCase and presentation and outputReport and busID and faultType):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
#                    raise ValueError('RevisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
#                    raise ValueError('ConfigName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputReport:
                    print('outputReport parameter is required. It should not be empty')
                if not busID:
                    print('busID parameter is required. It should not be empty')
                if not faultType:
                    print('faultType parameter is required. It should not be empty')
#                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()
    
    # AF endpoint is removed for Pepper due to Ui changes. Will be added in next release.
    # def runAF(self, revisionName: str, configName: str, studyCase: str, presentation: str, outputReport: str, getOnlineData: bool, onlineConfigOnly: bool, studyType: str, whatIfCommands=None) -> str:
    #     """Runs an Arc Flash study in ETAP.  Returns the output report location.  The results are available on the
    #        one-line diagram and in the output database (not available in Python).
    #         1. 'revisionName'              : The revision name for the study
    #         2. 'configName'                : The configuration name for the study
    #         3. 'studyCase'                 : The study case for the study
    #         4. 'presentation'              : The presentation to run the study in
    #         5. 'outputReport               : The output report name for the study results
    #         6. 'getOnlineData'             : Whether or not to get online data for the study
    #         7. 'onlineConfigOnly'          : Whether or not to get online configurations for the study
    #         8. 'studyType'                 : e.g., "ANSI ARC FLASH", "IEC ARC FLASH", "ANSI HV ARC FLASH 3P", "ANSI HV ARC FLASH LL", "ANSI HV ARC FLASH LG", "ANSI ARC FLASH 1PHASE", "IEC ARC FLASH 1PHASE"
    #         9. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')"""
            

    #     httpLocation = f'{self._baseAddress}{apiconstants.studies_runaf}?'

    #     params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}&studyType={7}".format(revisionName, configName, studyCase, presentation,
    #                                           outputReport, getOnlineData, onlineConfigOnly, studyType)
    #     # url encode params before sending
    #     address = httpLocation + urllib.parse.quote(params, safe='&=')
        
        

    #     try: 
    #         if (revisionName and configName and studyCase and presentation and outputReport and studyType):
    #             try:
    #                 result = datahub.post(address, whatIfCommands)
    #                 return result
                    
    #             except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
    #                 # handle error here
    #                 return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
    #         if not (revisionName and configName and studyCase and presentation and outputReport and studyType):
    #             if not revisionName:
    #                 print('revisionName parameter is required. It should not be empty')
    #             if not configName:
    #                 print('configName parameter is required. It should not be empty')
    #             if not studyCase:
    #                 print('studyCase parameter is required. It should not be empty')
    #             if not presentation:
    #                 print('presentation parameter is required. It should not be empty')
    #             if not outputReport:
    #                 print('outputFile parameter is required. It should not be empty')
    #             if getOnlineData is None:
    #                 print('getOnlineData parameter is required. It should not be empty')
    #             if onlineConfigOnly is None:
    #                 print('onlineConfigOnly parameter is required. It should not be empty')
    #             if not studyType:
    #                 print('studyType parameter is required. It should not be empty')
    #             raise ValueError('There are empty strings in one or more args! Please see the details above.')
    #     except ValueError as e:
    #         traceback.print_exc()

    def runCA(self, revisionName: str, configName: str, studyCase: str, presentation: str, 
              outputReport: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a CA study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation to run the study in
        5. 'outputReport'              : The output file for the study
        6. 'getOnlineData              : Whether or not to Get Online Data for the study
        7. 'onlineConfigOnly'          : Whether or not to Get online configurations for the study
        8. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')
        """
    
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runca}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData, onlineConfigOnly)
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        try: 
            if (revisionName and configName and studyCase and presentation and outputReport):
                try:
                    result = datahub.post(address, whatIfCommands)
                    return result
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputReport):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputReport:
                    print('outputReport parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()
            
    def runSO(self, revisionName: str, configName: str, studyCase: str, presentation: str, 
              outputReport: str, getOnlineData=None, onlineConfigOnly=None, whatIfCommands=None) -> str:
        """Runs a SO study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation to run the study in
        5. 'outputReport'              : The output file for the study
        6. 'getOnlineData              : Whether or not to Get Online Data for the study
        7. 'onlineConfigOnly'          : Whether or not to Get online configurations for the study
        8. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')
        """
    
        httpLocation = f'{self._baseAddress}{apiconstants.studies_runso}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData, onlineConfigOnly)
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')

        try: 
            if (revisionName and configName and studyCase and presentation and outputReport):
                try:
                    result = datahub.post(address, whatIfCommands)
                    return result
                except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
                    # handle error here
                    return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}
            if not (revisionName and configName and studyCase and presentation and outputReport):
                if not revisionName:
                    print('revisionName parameter is required. It should not be empty')
                if not configName:
                    print('configName parameter is required. It should not be empty')
                if not studyCase:
                    print('studyCase parameter is required. It should not be empty')
                if not presentation:
                    print('presentation parameter is required. It should not be empty')
                if not outputReport:
                    print('outputReport parameter is required. It should not be empty')
                raise ValueError('There are empty strings in one or more args! Please see the details above.')
        except ValueError as e:
            traceback.print_exc()

    def runUBSC(self, revisionName: str, configName: str, studyCase: str, presentation: str,
              outputReport: str, studyType: str, getOnlineData=None, onlineConfigOnly= None, whatIfCommands=None) -> str:
        """Runs a Unabalanced short circuit study in ETAP.  Returns the report name.  The results are available on
        the one-line diagram and in the output database (not available in Python).
        1. 'revisionName'              : The revision for the study
        2. 'configName'                : The config for the study
        3. 'studyCase'                 : The study case for the study
        4. 'presentation'              : The presentation for the study
        5. 'outputReport'              : The output file for the study
        6. 'getOnlineData              : Whether or not to Get Online Data for the study
        7. 'studyType'                 : Which HA analysis to run
        8. 'whatIfCommands'            : Python dictionary with what-if commands (e.g., 'OPEN TIE A')
        """

        httpLocation = f'{self._baseAddress}{apiconstants.studies_runubsc}?'
        params = "revisionName={0}&configName={1}&studyCase={2}&presentation={3}&outputReport={4}&getOnlineData={5}&onlineConfigOnly={6}&studyType={7}".format(
            revisionName, configName, studyCase, presentation, outputReport, getOnlineData,onlineConfigOnly,studyType)

        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params, safe='&=')
        
        result = datahub.post(address, whatIfCommands)

        # ------------------------------------------------
        # TODO 8/26/20: the response above is now JSON and no longer XML.
        # The code above produces this JSON:
        #
        # {
        #   "ReportPath": "E:\\FG1p5-Rel\\Example-ANSI\\Untitled.UL1S"
        # }
        #
        # The code below should download the output report when working from
        # a remote machine
        # ------------------------------------------------
        
        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine.')
            # TODO: -Neetin Choudary 2/16/2022
            # ipAddress is not passed and need to be tested for remote python api. This effects multiple studies.
            # Current scenario assumes base address as localhost only
            # print('Unable to download file from the remote machine ({}).'.format(
            #     ipAddressOrComputerName))
        except:
            print('Error: could not run HA (python-side exception).')


    def runStarZ(self, revisionName: str, configName: str, studyCase: str, outputReport: str,  whatIfCommands=None) -> str:
        """Runs a starZ study in ETAP.  Returns the report path.  The results are available on the
            one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'outputReport'              : The output file for the study
            5. 'whatIfCommands'            : Dictionary of what-if commands
        """

        # httpLocation = f'{self._baseAddress}{apiconstants.studies_runstarz}/'
        # params = "{0}/{1}/{2}/{3}".format(revisionName, configName, studyCase, outputReport)
        # address = httpLocation + urllib.parse.quote(params)

        httpLocation = f'{self._baseAddress}{apiconstants.studies_runstarz}?'
        
        param_list = [revisionName, configName, studyCase, outputReport]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"revisionName={param_list[0]}&configName={param_list[1]}&studyCase={param_list[2]}&outputReport={param_list[3]}"
        address = httpLocation + params


        result = datahub.post(address, whatIfCommands, token=self._token)
        
        
        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # P_Dict = json.loads(filename)  # Python Dict
            P_Dict = json.loads(result)  # Python Dict
            #print(P_Dict)
            P_Value = P_Dict["ReportPath"]
            
            a = '<string xmlns="http://schemas.microsoft.com/2003/10/Serialization/">'
            b = '</string>'
            result_fin = a + P_Value + b


            # Get path of the report
            report_path = ET.fromstring(result_fin).text
        
            return report_path
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     #print('Unable to download file from the remote machine ({}).'.format(ipAddressOrComputerName))
        #     print('Unable to download file from the remote machine')
        except:
            print('Error: could not run StarZ (python-side exception).')


    def runStarZasync(self, revisionName: str, configName: str, studyCase: str, outputReport: str,  whatIfCommands=None) -> str:
        """Runs a starZ study in ETAP.  Returns the report path.  The results are available on the
            one-line diagram and in the output database (not available in Python).
            1. 'revisionName'              : The revision for the study
            2. 'configName'                : The config for the study
            3. 'studyCase'                 : The study case for the study
            4. 'outputReport'              : The output file for the study
            5. 'whatIfCommands'            : Dictionary of what-if commands
        """

        # httpLocation = f'{self._baseAddress}{apiconstants.studies_runstarzasync}/'
        # params = "{0}/{1}/{2}/{3}".format(revisionName, configName, studyCase, outputReport)
        # address = httpLocation + urllib.parse.quote(params)


        httpLocation = f'{self._baseAddress}{apiconstants.studies_runstarzasync}?'
        
        param_list = [revisionName, configName, studyCase, outputReport]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"revisionName={param_list[0]}&configName={param_list[1]}&studyCase={param_list[2]}&outputReport={param_list[3]}"
        address = httpLocation + params


        result = datahub.post(address, whatIfCommands, token=self._token)


        try:
            # filename = self.retrieveDownloadedFilePath(result)
            # return filename
            return result
        except KeyboardInterrupt:
            print('Operation cancelled.')
        # except FileNotFoundError:
        #     print('Unable to download file from the remote machine ({}).'.format(
        #         ipAddressOrComputerName))
        except:
            print('Error: could not run TDLF (python-side exception).')


            
    def downloadRemoteFile(self, folderName, resultFile: str):
        """Download results file from the remote ETAP machine. 
        1. 'foldername'              : Directory where the file gets downloaded
        2. 'resultFile'              : Path on the remote ETAP machine where the results file resides"""

        # encode the file path string as it can contain invalid path characters and it might throw
        encodedBytes = base64.b64encode(resultFile.encode("utf-8"))
        encodedStr = str(encodedBytes, "utf-8")

        httpLocation = f'{self._baseAddress}{apiconstants.application_downloadfile}/'
        params = "{0}".format(encodedStr)
        # httpLocation = "http://{0}:{1}/pythonservice2/{2}/downloadfile/{3}".format(
        #     ipAddressOrComputerName, portNumber, projectNameNoExtension, encodedStr)
        # test file download url
        # httpLocation = "http://www.ovh.net/files/10Mb.dat"

        address = httpLocation + urllib.parse.quote(params)
        filename = os.path.join(folderName, os.path.basename(resultFile))
        datahub.get_file_progressbar(address, filename, token=self._token)

    def getFilename_fromCD(self, cd: str):
        """
        Get filename from content-disposition
        """
        if not cd:
            return None
        fname = re.findall('filename=(.+)', cd)
        if len(fname) == 0:
            return None
        return fname[0]

    def chunk_report(self, bytes_so_far, chunk_size, total_size):
        """
        Report download progress
        """
        percent = float(bytes_so_far) / total_size
        percent = round(percent*100, 2)
        # sys.stdout.write("Downloaded %d of %d bytes (%0.2f%%)\r" %
        #    (bytes_so_far, total_size, percent))
        # -----
        #roundedPercent = int(100*bytes_so_far/total_size)
        #sys.stdout.write('\r[{}{}] {}% completed'.format('█' * roundedPercent, '.' * (100-roundedPercent), roundedPercent))
        self.showProgressBar(bytes_so_far, total_size)

        if bytes_so_far >= total_size:
            sys.stdout.write('\n')

    def chunk_read(self, response, chunk_size=8192, report_hook=None):
        """
        Read chunks of bytes and report progress
        """
        total_size = response.headers.get('content-length').strip()
        total_size = int(total_size)
        bytes_so_far = 0
        data = b''
        chunk_size = max(int(total_size/1000), 1024*1024)
        while 1:
            chunk = response.read(chunk_size)
            bytes_so_far += len(chunk)
            data += chunk
            if not chunk:
                break

            if report_hook:
                report_hook(bytes_so_far, chunk_size, total_size)

        return data

    def showProgressBar(self, bytes_so_far, total_size):
        """
        Calculate progress % and display progress bar
        """
        roundedPercent = int(100*bytes_so_far/total_size)
        sys.stdout.write('\r[{}{}] {}% completed'.format(
            '█' * roundedPercent, '.' * (100-roundedPercent), roundedPercent))

    def __removeUnwantedChars(self, xmlString):
        """
        Replace the unwanted chars (ÿþ) added by DataHub Python for Unicode encoding
        """
        if xmlString.startswith('ÿþ'):
            # return xmlString.replace('ÿþ','')
            return xmlString[len('ÿþ'):]
        return xmlString

    def __removeXmlVersionTag(self, xmlString):
        """
        Remove the <?xml version="1.0" encoding="utf-16"?> tag added by DataHub Python for Unicode encoding
        """
        if xmlString.find('<?xml version="1.0" encoding="utf-16"?>') >= 0:
            return xmlString.replace('<?xml version="1.0" encoding="utf-16"?>', '')
        return xmlString

    def __decodeUnicodeEncodedString(self, xmlString):
        """
        Decode the Unicode encoded string
        """
        return xmlString.encode().decode('utf-16')

    def retrieveDownloadedFilePath(self, result) -> str:
        """
        Returns the complete file path of the downloaded file
        1. 'result'              : The complete file path on the remote system
        """

        if self._isForRemoteEtap == True:
                if result:
                    jsonOutputPath = json.loads(result)
                    folder = os.path.realpath(tempfile.gettempdir())
                    self.downloadRemoteFile(folder, jsonOutputPath["ReportPath"])
                    filename = os.path.join(folder, os.path.basename(jsonOutputPath["ReportPath"]))
                    # print("File downloaded successfully at the following location: {}".format(filename))
                    jsonOutputPath["ReportPath"] = filename
                    return json.dumps(jsonOutputPath,
                                    indent=4, sort_keys=False,
                                    separators=(',', ':'), ensure_ascii=False)
                else:
                    raise FileNotFoundError
        return result
    
