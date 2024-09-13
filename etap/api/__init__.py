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

# hoist classes one-level up so that one can type s = etap.studies()
#from .etap_client import EtapClient
from etap.api.other.etap_client import EtapClient

def connect(baseAddress:str, projectName=None)->EtapClient:
    """Forms a connection with ETAP.  This method should be called before making any
    other call to ETAP.

    Args:
        baseAddress (str): DataHub base address (e.g., 'http://localhost:50000')

    Returns:
        EtapClient: An instance of the client used to communicate with an instance of ETAP
    """    
    e = EtapClient()
    e.connect(baseAddress, projectName)
    
    return e
    #self._ipAddressOrComputerName = ipAddressOrComputerName
    # self._portNumber = portNumber + 2 # true port number is the DataHub port number + 2
    #self._projectNameNoExtension = re.sub(r'[^a-zA-Z0-9_]', 'x', projectNameNoExtension)
    #self._isForRemoteEtap = self.__isRequestForRemoteMachine(ipAddressOrComputerName)


def getVersion() -> str:
    """Returns the ETAP package version.  This package is the entry point to the ETAP Python API."""
    return "0.2.0"

# [PrashanthM]++ 2020/01/22 IR-67502 - Develop Python Interface using xml string for fault location & run StarZ
# import eProtect module to run studies: Relay setting assistant, Fault location, Signal injection
#from etap.eProtect import eProtect

# import Units enum for eProtect
#from etap.eProtect import Units


def printDict(dict):
    """Prints a dictionary as a nice string."""
    print("{")
    for key in dict:
        print("    \"{0}\": \"{1}\"".format(key, dict[key]))
    print("}")
