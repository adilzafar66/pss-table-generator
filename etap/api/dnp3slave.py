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


from .other import datahub as datahub
from .other import api_constants as apiconstants
import requests

class Dnp3Slave:

    def __init__(self, baseAddress:str, sentToken:str, projectName=None):
        self._baseAddress = baseAddress
        self._token = sentToken
        self._projectName = projectName

    def getsampleaidiruntimejson(self) -> str:
        """Returns sample runtime JSON for updating analog input and digital input values of the DNP3 slave"""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getsampleaidiruntimejson}', token=self._token)

    def getsampleaodoruntimejson(self) -> str:
        """Returns sample runtime JSON for updating analog output and digital output values of the DNP3 slave"""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getsampleaodoruntimejson}', token=self._token)

    def getsampleinitjson(self) -> str:
        """Return sample init JSON for starting the DNP3 slave"""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getsampleinitjson}', token=self._token)

    def getslaveai(self) -> str:
        """Return analog input of the DNP3 slave."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslaveai}', token=self._token)

    def getslaveao(self) -> str:
        """Return analog output of the DNP3 slave.."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslaveao}', token=self._token)


    def getslavechannelstate(self) -> str:
        """Returns channel state of the DNP3 slave. Expected states are below:
            NULL: Slave is not running
            CLOSED: Master is disconnected
            OPENING: Slave is started. Waiting for Master to connect
            OPEN: Master is connected to slave
            SHUTDOWN: Slave is stopped and disabled until next start"""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslavechannelstate}', token=self._token)


    def getslavedb(self) -> str:
        """Return all DNP3 slave datapoints."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslavedb}', token=self._token)

    def getslavedi(self) -> str:
        """Return digital input of the DNP3 slave."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslavedi}', token=self._token)

    def getslavedo(self) -> str:
        """Return digital output of the DNP3 slave."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslavedo}', token=self._token)

    def getslavesettings(self) -> str:
        """Return all settings of the DNP3 slave."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslavesettings}', token=self._token)
        

#    def getsampleruntimejson(self) -> str:
#        """Return sample runtime JSON for updating analog input and digital input values of the DNP3 slave"""
#        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getsampleruntimejson}')       
    
    def getslavestate(self) -> str:
        """Return state of the DNP3 slave."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_getslavestate}', token=self._token) 

    def dnp3ping(self) -> str:
        """Returns true if the host machine is reachable."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_dnp3ping}', token=self._token)

    def start(self, dnp3SettingsInit=None) -> str:
        """Initialize and start the DNP3 slave using the given DNP3 settings data in JSON format."""
        httpLocation = f'{self._baseAddress}{apiconstants.dnp3slave_start}'
        address = httpLocation

        try:
            result = datahub.post(address, dnp3SettingsInit, token=self._token)
            return result
#            try:
#                filename = self.retrieveDownloadedFilePath(result)
#                return filename                  
#            except KeyboardInterrupt:
#                print('Operation cancelled.')
#            except FileNotFoundError:
#                print('Unable to download file from the remote machine ({}).'.format(
#                    ipAddressOrComputerName))
#            except:
#                print('Error: could not initialize and start the DNP3 slave using the given DNP3 settings data in JSON format.')
        except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
            # handle error here
            return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}


    def stop(self) -> str:
        """Return true if the DNP3 is stopped."""
        return datahub.get(f'{self._baseAddress}{apiconstants.dnp3slave_stop}', token=self._token)


    def updateaidi(self, dnp3SettingsRuntime=None) -> str:
        """Updates the DNP3 slave using the given DNP3 settings data in JSON format. NOTE: to get real sample data use the / api / v1 / getsampleaidiruntimejson endpoint"""
        httpLocation = f'{self._baseAddress}{apiconstants.dnp3slave_updateaidi}'
        address = httpLocation

        try:
            result = datahub.post(address, dnp3SettingsRuntime, token=self._token)
            return result
#            try:
#                filename = self.retrieveDownloadedFilePath(result)
#                return filename                  
#            except KeyboardInterrupt:
#                print('Operation cancelled.')
#            except FileNotFoundError:
#                print('Unable to download file from the remote machine ({}).'.format(
#                    ipAddressOrComputerName))
#            except:
#                print('Error: could not update the DNP3 slave using the given DNP3 settings data in JSON format.')
        except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
            # handle error here
            return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}


    def updateaodo(self, dnp3SettingsRuntime=None) -> str:
        """Updates the DNP3 slave using the given DNP3 settings data in JSON format. NOTE: to get real sample data use the / api / v1 / getsampleaodoruntimejson endpoint"""
        httpLocation = f'{self._baseAddress}{apiconstants.dnp3slave_updateaodo}'
        address = httpLocation

        try:
            result = datahub.post(address, dnp3SettingsRuntime, token=self._token)
            return result
#            try:
#                filename = self.retrieveDownloadedFilePath(result)
#                return filename                  
#            except KeyboardInterrupt:
#                print('Operation cancelled.')
#            except FileNotFoundError:
#                print('Unable to download file from the remote machine ({}).'.format(
#                    ipAddressOrComputerName))
#            except:
#                print('Error: could not update the DNP3 slave using the given DNP3 settings data in JSON format.')
        except (requests.exceptions.HTTPError, ValueError, NameError) as exception:
            # handle error here
            return {'Response Body': exception.response.text, 'Response Code': exception.response.status_code, 'Response Headers': exception.response.headers}