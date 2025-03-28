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





#This modules makes it easy to produce GET and POST calls using an underlying
#third-part http library.  For example, this module
#may wrap 'urllib', 'urllib2', or the 'requests' package.
#This wrapper also isolates etapPy from the underlying http-request library.

import base64
import json
import os
import re
import signal
import subprocess
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime
import requests
requests.packages.urllib3.disable_warnings()
from requests.exceptions import RequestException
import sys

from .. import settings

__version__ = "24.0.0"

def get(urlAbsolute: str, token=None) -> str:
    """Makes an http GET call against ETAP DataHub using the given absolute URL
    Args:
        urlAbsolute (str):  Absolute URL
    Returns:
        str: JSON response from the http GET call (if any)
    """
    

    # if url.startswith('/') is False:
    #    url = '/' + url
    if token:
        resp = requests.request("GET", urlAbsolute, verify=not(settings.https_enable), headers={"Authorization": token})
    else:
        resp = requests.request("GET", urlAbsolute, verify=not(settings.https_enable))
    
    # resp.raise_for_status()
    return resp.text


def get_file(urlAbsolute:str, writeFullPath:str, token=None):
    """Downloads a file from ETAP.  The file is written locally to disk to
    the location specified in the arguments.
    Args:
        urlAbsolute (str):  Absolute URL (e.g., http://localhost:port/etap/api/v1/application/downloadfile/{fullPathToFileAsBase64})
                            where fullPathToFileAsBase64 is the full path to the file 
                            on the ETAP machine as a base64-encoded string
        writeFullPath (str):     Full path to where file should be written to
    """
    if token:
        resp = requests.get(urlAbsolute, allow_redirects=True, verify=not(settings.https_enable), headers={"Authorization": token})
    else:
        resp = requests.get(urlAbsolute, allow_redirects=True, verify=not(settings.https_enable))
        
    # resp.raise_for_status()
    f = open(writeFullPath, 'wb')
    f.write(resp.content)
    f.close()


def get_file_progressbar(urlAbsolute:str, writeFullPath:str, token=None):
    """Downloads a file from ETAP and displays progress.  The file is written locally to disk to
    the location specified in the arguments.
    Args:
        urlAbsolute (str):  Absolute URL (e.g., http://localhost:port/etap/api/v1/application/downloadfile/{fullPathToFileAsBase64})
                            where fullPathToFileAsBase64 is the full path to the file 
                            on the ETAP machine as a base64-encoded string
        writeFullPath (str):     Full path to where file should be written to
    """
    
    with open(writeFullPath, "wb") as f:
        print("Downloading %s" % os.path.basename(writeFullPath))
        if token:
            response = requests.get(urlAbsolute, stream=True, allow_redirects=True, verify=not(settings.https_enable), headers={"Authorization": token})
        else:
            response = requests.get(urlAbsolute, stream=True, allow_redirects=True, verify=not(settings.https_enable))
        # response.raise_for_status()
        total_length = response.headers.get('content-length')

        if total_length is None: # no content length header
            f.write(response.content)
        else:
            dl = 0
            total_length = int(total_length)
            for data in response.iter_content(chunk_size=4096):
                dl += len(data)
                f.write(data)
                done = int(50 * dl / total_length)
                #sys.stdout.write("\r[%s%s]" % ('=' * done, ' ' * (50-done)) )    
                roundedPercent = int(100*dl/total_length)
                sys.stdout.write('\r[{}>{}] {}% completed'.format(
                '=' * roundedPercent, ' ' * (100-roundedPercent), roundedPercent))
                sys.stdout.flush()


def post(urlAbsolute: str, someDict, token=None, headers=None) -> str:
    """Makes an http POST call against ETAP DataHub using the given absolute URL
    Args:
        urlAbsolute (str):  Absolute URL
        someDict (dict): Dictionary with content to post in the body
    Returns:
        str: JSON response from the http POST call (if any)
    """

    if token:
        resp = requests.post(urlAbsolute, json=someDict, verify=not(settings.https_enable), headers={"Authorization": token})
    else:
        if headers:
            resp = requests.post(urlAbsolute, json=someDict, verify=not(settings.https_enable), headers=headers)
        else:
            resp = requests.post(urlAbsolute, json=someDict, verify=not(settings.https_enable))
            
    # resp.raise_for_status()
    return resp.text




class DataHubClient():
    """Client to communicate with a particular instance of ETAP DataHub
    """

    def __init__(self, ipAddressOrComputerName: str, portNumber: int):
        """Initializer function.

        Args:
            ipAddressOrComputerName (str): DataHub IP address or hostname
            portNumber (int): DataHub port number
        """        
        self.baseAddress = f"https://{ipAddressOrComputerName}:{str(portNumber)}"

    def httpGet(self, relativeUrl: str) -> str:
        """Makes an http GET call against ETAP DataHub using the relative URL
        Args:
            relativeUrl (str):  URL relative to base address (e.g., '/etap/api/v1')
        Returns:
            str: Response from http GET call (if any)
        """
        if relativeUrl.startswith('/') is False:
            raise ValueError("### Relative URL must start with '/' character")

        relativeUrl = f"{self.baseAddress}{relativeUrl}"
        resp = requests.request("GET", relativeUrl, verify=not(settings.https_enable))
        # resp.raise_for_status()
        return resp.text

    def getBaseAddress(self) -> str:
        """Returns the base address used by this DataHubClient instance

        Returns:
            str: base address in the form 'http://host:port'
        """
        return self.baseAddress
