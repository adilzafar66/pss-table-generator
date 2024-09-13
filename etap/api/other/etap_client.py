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
import sys
import urllib.request
import re
import json
import xml.etree.ElementTree as ET
import urllib.parse
import base64
import os
import subprocess
import signal
from pathlib import Path

from .. import studies as studies
from .. import application as application
from .. import projectdata as projectdata
from .. import scenario as scenario
from .. import dnp3slave as dnp3slave
from netifaces import interfaces, ifaddresses, AF_INET
import socket
from . import datahub


class EtapClient:
    """Client connection class used to communicate with ETAP.  A connection must be established before interacting with ETAP."""
    _etapPath = None
    
    def connect(self, baseAddress, projectName=None):
        self._baseAddress = baseAddress
        self._projectName = projectName
        self._token = self.getToken()
        self.application  = application.Application(baseAddress, self._token, self._projectName)
        self.projectdata  = projectdata.ProjectData(baseAddress, self._token, self._projectName)
        self.scenario  = scenario.Scenario(baseAddress, self._token, self._projectName)        
        self.scenario_WhatifCommands  = scenario.Scenario.WhatifCommands(baseAddress, self._token, self._projectName)   
        self._ipAddress = self.__extractIpAddressFromBaseAddress(baseAddress)
        self._isForRemoteEtap = self.__isRequestForRemoteMachine(self._ipAddress)
        self.studies  = studies.Studies(baseAddress, self._isForRemoteEtap, self._token, self._projectName)
        self.dnp3slave  = dnp3slave.Dnp3Slave(baseAddress, self._token, self._projectName)   

    # def connect(self, ipAddressOrComputerName, portNumber):
    #     """Forms a connection with ETAP.  This method should be called before making any
    #     other call to ETAP."""
    #     self._ipAddressOrComputerName = ipAddressOrComputerName
    #     self._portNumber = portNumber + 2 # true port number is the DataHub port number + 2
    #     #self._projectNameNoExtension = re.sub(r'[^a-zA-Z0-9_]', 'x', projectNameNoExtension)
    #     self._isForRemoteEtap = self.__isRequestForRemoteMachine(ipAddressOrComputerName)
    #     #print(f"Connected to ETAP @ http://{ipAddressOrComputerName}:{str(portNumber)}")

    
    def getToken(self):
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            app_path = Path(sys._MEIPASS)
        else:
            app_path = Path(sys.argv[0]).parent
        with open(app_path / 'res' / 'keyFile.json', 'r') as f:
            content_list = json.load(f)
        for dicts in content_list:
            if dicts["projectName"] == self._projectName:
                return dicts["token"]
        self._token = None
    

    def getProcessId(self)->int:
        """Returns the Process Id of the etap client"""

        address = "http://{}:{}/pythonservice2/{}/pid".format(self._ipAddressOrComputerName, self._portNumber, self._projectNameNoExtension)
        try:
            h = urllib.request.urlopen(address)
            xmlString = "".join(map(chr, h.read()))
            xmlString = self.__removeUnwantedChars(xmlString)
            xmlString = self.__decodeUnicodeEncodedString(xmlString)
            xmlString = self.__removeXmlVersionTag(xmlString)
            rtn = xmlString[xmlString.find(">") + 1:xmlString.rfind("<")]
            return rtn
        except :
            return -1

    def close(self):
        """Kills the connected instance of etap"""

        pid = int(self.getProcessId())
        if(pid > -1):
            os.kill(pid, signal.SIGTERM)
            
    def setEtapPath(path):
        EtapClient._etapPath = path
        
    def open(project = None):
        if(EtapClient._etapPath == None):
            print("Please set the path of etaps64.exe with setEtapPath first")
        else:
            return subprocess.Popen(r"{}\etaps64.exe {}".format(EtapClient._etapPath, project))
            return False

    
    def __extractIpAddressFromBaseAddress(self, baseAddress):
        """Extracts IP address from the base address"""
        regex = re.compile('{}(.*){}'.format(re.escape('//'), re.escape(':')))
        ipAddresses = regex.findall(baseAddress)
        return ipAddresses[0]

    def __isRequestForRemoteMachine(self, ipAddress):
        """Determines whether this request if for a remote machine running ETAP"""
        result = True
        ip_list = []    
        try:
            for interface in interfaces():            
                if len(ifaddresses(interface)) - 1 < AF_INET:
                    continue
                for link in ifaddresses(interface)[AF_INET]:
                    ip_list.append(link['addr'])
            index = ip_list.index(ipAddress)
            if index >= 0:
                result = False
            return result
        except ValueError:
            return result
        
        # return socket.gethostbyname(socket.gethostname()) != ipAddress


    def __removeUnwantedChars(self, xmlString):
        """
        Replace the unwanted chars (ÿþ) added by DataHub Python for Unicode encoding
        """
        if xmlString.startswith('ÿþ'):
            #return xmlString.replace('ÿþ','')
            return xmlString[len('ÿþ'):]
        return xmlString


    def __removeXmlVersionTag(self, xmlString):
        """
        Remove the <?xml version="1.0" encoding="utf-16"?> tag added by DataHub Python for Unicode encoding
        """
        if xmlString.find('<?xml version="1.0" encoding="utf-16"?>') >= 0:
            return xmlString.replace('<?xml version="1.0" encoding="utf-16"?>','')
        return xmlString


    def __decodeUnicodeEncodedString(self, xmlString):
        """
        Decode the Unicode encoded string
        """            
        return xmlString.encode().decode('utf-16')
