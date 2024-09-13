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

import os
import json
from .other import datahub as datahub
from .other import api_constants as apiconstants
import urllib.request
import urllib.parse

class Application:

    def __init__(self, baseAddress:str, sentToken:str, projectName=None):
        self._baseAddress = baseAddress
        self._token = sentToken
        self._projectName = projectName
        
    
    def downloadfile(self, filePathAsBase64: str, writeFullPath:str):
        """Download a file from ETAP.
        
        ArgsL
           writeFullPath (str):     Full path to where file should be written to"""
        # return datahub.get_file(f'{self._baseAddress}{apiconstants.application_downloadfile}/{filePathAsBase64}', writeFullPath, token=self._token)

        httpLocation = f'{self._baseAddress}{apiconstants.application_downloadfile}?fullPathToFileAsBase64='
        params = f"{filePathAsBase64}"
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params)

        return datahub.get_file(address, writeFullPath, token=self._token)
    
    def filepaths(self) -> str:
        """Return filepath information about ETAP."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_filepaths}', token=self._token)

    def pid(self) -> str:
        """Return the process ID of the ETAP instance."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_pid}', token=self._token)

    def ping(self) -> bool:
        """Return True if ETAP can be reached; False otherwise."""
        try:
            response = datahub.get(f'{self._baseAddress}{apiconstants.application_ping}', token=self._token)
            return "true" in response
        except:
            return False

    def projectfile(self) -> str:
        """Return information about the project currently open."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_projectfile}', token=self._token)

    def showhelp(self, section:str) -> str:
        """Open the ETAP help file to the PSCAD section."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_showhelp}/{section}', token=self._token)

    def version(self) -> str:
        """Return the ETAP version."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_version}', token=self._token)

    def getcurrentuser(self) -> str:
        """Return current user name."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getcurrentuser}', token=self._token)

    def getallusers(self) -> str:
        """Return all user names."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getallusers}', token=self._token)

    def getlanguage(self) -> str:
        """Return language string."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getlanguage}', token=self._token)

    def getcurrentzoomlevel(self) -> str:
        """Return current zoom level."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getcurrentzoomlevel}', token=self._token)

    def getactivescenario(self) -> str:
        """Return active scenario."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getactivescenario}', token=self._token)        

    def getmessagelog(self) -> str:
        """Return ETAP message log."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getmessagelog}', token=self._token)      

    def clearmessagelog(self) -> str:
        """Clear ETAP message log."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_clearmessagelog}', token=self._token)   

    def hideetapapplication(self) -> str:
        """Hide or minimize ETAP application.""" 
        httpLocation = f'{self._baseAddress}{apiconstants.application_hideetapapplication}'
        emptyDict = dict()
        address = httpLocation
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    def showetapapplication(self) -> str:
        """Show or restore ETAP application.""" 
        httpLocation = f'{self._baseAddress}{apiconstants.application_showetapapplication}'
        emptyDict = dict()
        address = httpLocation
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    def hidemessagelog(self) -> str:
        """Hide message log.""" 
        httpLocation = f'{self._baseAddress}{apiconstants.application_hidemessagelog}'
        emptyDict = dict()
        address = httpLocation
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    def showmessagelog(self) -> str:
        """Show message log.""" 
        httpLocation = f'{self._baseAddress}{apiconstants.application_showmessagelog}'
        emptyDict = dict()
        address = httpLocation
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    def savemessagelog(self, filePathAsBase64: str) -> str:
        """Save ETAP message log.""" 
        # httpLocation = f'{self._baseAddress}{apiconstants.application_savemessagelog}/{filePathAsBase64}'
        httpLocation = f'{self._baseAddress}{apiconstants.application_savemessagelog}?fullPathToFileAsBase64='
        params = filePathAsBase64
        emptyDict = dict()
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.post(address, emptyDict, token=self._token)
        return result
    
    def getstudymodes(self) -> str:
        """Return ETAP study modes."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_getstudymodes}', token=self._token)   
    
    def setstudymode(self, study_mode: str) -> str:
        """Set ETAP study modes."""

        httpLocation = f'{self._baseAddress}{apiconstants.application_setstudymode}?studyMode='
        params = study_mode
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params)
        
        emptyDict = dict()
        return datahub.post(address,emptyDict, token=self._token) 
    
    def login(self, token):
        """Return True if login successful; False otherwise."""
        try:
            # httpLocation = f'{self._baseAddress}{apiconstants.application_login}/{token}'
            httpLocation = f'{self._baseAddress}{apiconstants.application_login}?token='
            params = token
            address = httpLocation + urllib.parse.quote(params)
            emptyDict = dict()
            # address = httpLocation
            response = datahub.post(address,emptyDict)
            response_dict = json.loads(response)
            returnedToken = response_dict['Value']
            
            with open(os.path.dirname(os.path.abspath(__file__)) + "\\other\\keyFile.json", 'r') as f:
                content_list = json.load(f)
            
            content_list.append({'projectName':self._projectName, 'token':returnedToken})
            
            with open(os.path.dirname(os.path.abspath(__file__)) + "\\other\\keyFile.json", 'w') as f:
                json.dump(content_list, f)
            
            return True
            
        except:
            return False
        
    def getstudytypes(self, study_mode: str) -> str:
        """Set ETAP study types."""
        # httpLocation = f'{self._baseAddress}{apiconstants.application_getstudytypes}/'
        # params = "{0}".format(study_mode)

        httpLocation = f'{self._baseAddress}{apiconstants.application_getstudytypes}?studyMode='
        params = study_mode
        
        # url encode params before sending
        address = httpLocation + urllib.parse.quote(params)

        result = datahub.get(address)
        return result

    def exitetap(self) -> str:
        """Exit ETAP gracefully."""
        return datahub.get(f'{self._baseAddress}{apiconstants.application_exitetap}', token=self._token)

    def validatetoken(self, datahub_token: str) -> str:
        """Validates datahub access token."""
        # httpLocation = f'{self._baseAddress}{apiconstants.application_validatetoken}/'
        # params = "{0}".format(datahub_token)
        httpLocation = f'{self._baseAddress}{apiconstants.application_validatetoken}?token='
        params = datahub_token
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result
