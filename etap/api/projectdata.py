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
import urllib.request
import urllib.parse

class ProjectData:

    def __init__(self, baseAddress:str, sentToken:str, projectName=None):
        self._baseAddress = baseAddress
        self._token = sentToken
        self._projectName = projectName


    def getallelementdata(self, elementType:str) -> str:
        """Returns XML for elements of a particular type."""
        # return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getallelementdata}/{elementType}', token=self._token)
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getallelementdata}?elementType='
        params = elementType
        address = httpLocation + urllib.parse.quote(params)
        # return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getallelementdata}/{elementType}', token=self._token)
        return datahub.get(address, token=self._token)

    def getbusnames(self) -> str:
        """Returns all bus names."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getbusnames}', token=self._token)

    def getconfigurations(self) -> str:
        """Returns all configuration names."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getconfigurations}', token=self._token)
    
    def getelementnames(self, elementType: str) -> str:
        """Returns element names by element type."""

        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getelementnames}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getelementnames}?elementType='
        # params = "{0}".format(elementType)
        params = elementType
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result

    def getrevisions(self) -> str:
        """Returns all revision names."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getrevisions}', token=self._token)

    def getstudymodesandcases(self) -> str:
        """Returns all study mode and study case names."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getstudymodesandcases}', token=self._token)
    
    def getxml(self) -> str:
        """Returns XML for the current project.  The first time this function is called
        the response takes longer than usual.  Response time may also be affected by model size."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getxml}', token=self._token)
    
    # def getxmlflat(self) -> str:
    #     """Returns Flat XML for the current project.  The first time this function is called
    #     the response takes longer than usual.  Response time may also be affected by model size."""
    #     return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getxmlflat}', token=self._token)

    def getelementprop(self, elementType: str, elementName: str, fieldName: str) -> str:
        """Get value of a field of the property of an element

        Args:
            elementType (str): The type of element
            elementName (str): The name of element
            fieldName (str): The name of the field of element property

        Returns:
            str: value or empty string '' for failure
        """        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getelementprop}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getelementprop}?elementType='
        param_list = [elementType, elementName, fieldName]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"{param_list[0]}&elementName={param_list[1]}&fieldName={param_list[2]}"
        # params = "{0}/{1}/{2}".format(elementType, elementName, fieldName)
        address = httpLocation + params
        result = datahub.get(address, token=self._token)
        return result

    def getelementpropertynamesxml(self, elementType: str):
        """Get supported element properties names in xml format for given element type
        Args:
            elementType (str): The type of element
        
        Returns:
            str: value or empty string '' for failure
        """
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getelementpropertynamesxml}/'
        # params = "{0}".format(elementType)
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getelementpropertynamesxml}?elementType='
        params = elementType
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result
    
    def getelementtypes(self) -> str:
        """Returns all element types supported."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getelementtypes}', token=self._token)
    
    def setelementprop(self, elementType: str, elementName: str, fieldName: str, value: str) -> str:
        """Set value of a field of an element property

        Args:
            elementType (str): The type of element
            elementName (str): The name of element
            fieldName (str): The name of the field of element property
            value (str): The value to be set

        Returns:
            str: 'True' for success or 'False' for failure
        """       
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_setelementprop}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_setelementprop}?elementType='

        param_list = [elementType, elementName, fieldName, value]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"{param_list[0]}&elementName={param_list[1]}&fieldName={param_list[2]}&value={param_list[3]}"
        # params = "{0}/{1}/{2}".format(elementType, elementName, fieldName)
        address = httpLocation + params

        # params = "{0}/{1}/{2}/{3}".format(elementType, elementName, fieldName, value)
        emptyDict = dict()
        # address = httpLocation + urllib.parse.quote(params)
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    # def sendpdexml(self, xmlElement) -> str:
    #     """Send XML to ETAP PDE

    #     Args:
    #         xmlElement: xml str

    #     Returns:
    #         str: 'True' for success
    #     """       
    #     httpLocation = f'{self._baseAddress}{apiconstants.projectdata_sendpdexml}'
    #     result = datahub.post(httpLocation, xmlElement, token=self._token)
    #     return result

    # def createelementinold(self, elementType: str, elementName: str, compositeNetwork=None, locationX=None, locationY=None) -> str:
    def createelementinold(self, elementType: str, elementName: str, compositeNetwork="", locationX="", locationY="") -> str:
        """Create an element in OLD

        Args:
            elementType (str): The type of element
            elementName (str): The name of element
            compositeNetwork (str): The name of the compositeNetwork
            locationX (str): The X position value to be set
            locationY (str): The Y position value to be set

        Returns:
            str: 'True' for success or 'False' for failure
        """       
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_createelementinold}/'
        # params = "elementType={0}&elementName={1}&compositeNetwork={2}&flat={3}&locationX={4}&locationY={5}".format(elementType, elementName, compositeNetwork, flat, locationX, locationY)
        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_createelementinold}?'
        # params = "elementType={0}&elementName={1}&compositeNetwork={2}&locationX={3}&locationY={4}".format(elementType, elementName, compositeNetwork, locationX, locationY)

        # # url encode params before sending
        # address = httpLocation + urllib.parse.quote(params, safe='&=')


        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_createelementinold}?'

        param_list = [elementType, elementName, compositeNetwork, locationX, locationY]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"elementType={param_list[0]}&elementName={param_list[1]}&compositeNetwork={param_list[2]}&locationX={param_list[3]}&locationY={param_list[4]}"
        address = httpLocation + params


        emptyDict = dict()
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    def deleteelementinold(self, elementType: str, elementName: str) -> str:
        """Delete an element in OLD

        Args:
            elementType (str): The type of element
            elementName (str): The name of element

        Returns:
            str: 'True' for success or 'False' for failure
        """       

        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_deleteelementinold}?elementType='

        param_list = [elementType, elementName]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"{param_list[0]}&elementName={param_list[1]}"
        address = httpLocation + params
                
        emptyDict = dict()
        result = datahub.post(address, emptyDict, token=self._token)
        return result

    def iselementenergized(self, elementName: str) -> str:
        """Get whether an element is energized or not

        Args:
            elementName (str): The name of element

        Returns:
            str: {"Value": "True"} for success
        """        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementenergized}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementenergized}?elementName='
        # params = "{0}".format(elementName)
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result
        

    def iselementhidden(self, elementName: str) -> str:
        """Get whether an element is hidden or not

        Args:
            elementName (str): The name of element

        Returns:
            str: {"Value": "True"} for success
        """        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementhidden}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementhidden}?elementName='
        # params = "{0}".format(elementName)
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result
        
    def iselementnode(self, elementName: str) -> str:
        """Get whether an element is a node or not

        Args:
            elementName (str): The name of element

        Returns:
            str: {"Value": "True"} for success
        """        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementnode}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementnode}?elementName='
        # params = "{0}".format(elementName)
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result
        
    def iselementoutofservice(self, elementName: str) -> str:
        """Get whether an element is out of service or not

        Args:
            elementName (str): The name of element

        Returns:
            str: {"Value": "True"} for success
        """        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementoutofservice}/'
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementoutofservice}?elementName='
        # params = "{0}".format(elementName)
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        result = datahub.get(address, token=self._token)
        return result

    def getstudycasenames(self) -> str:
        """Returns a list of study case names (IDs).."""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getstudycasenames}', token=self._token)
    
    def getstudymodesandcases(self) -> str:
        """Returns a dictionary of study modes and their corresponding study cases"""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getstudymodesandcases}', token=self._token)

    def getstudycase(self, studyCaseId:str) -> str:
        """Returns a study case by ID."""
        # return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getstudycase}/{studyCaseId}', token=self._token)
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getstudycase}?studyCaseId='
        params = studyCaseId
        address = httpLocation + urllib.parse.quote(params)
        # return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getstudycase}/{studyCaseId}', token=self._token)
        return datahub.get(address, token=self._token)

    def setstudycase(self, xmlElement) -> str:
        """Adds a new (or replaces an existing) study case."""       
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_setstudycase}'
        result = datahub.post(httpLocation, xmlElement, token=self._token)
        return result
    
    def validatecredentials(self, authType: str, userName: str, password: str) -> str:
        """Validate credentials for EtapAuthentication and WindowsAuthentication

        Args:
            authType (str): The type of authentication
            userName (str): The name of the user
            password (str): The password for the userName

        Returns:
            str: value or empty string '' for failure
        """        
        # httpLocation = f'{self._baseAddress}{apiconstants.projectdata_validatecredentials}/'
        # params = "{0}/{1}/{2}".format(authType, userName, password)
        # address = httpLocation + urllib.parse.quote(params)


        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_validatecredentials}?'
        
        param_list = [authType, userName, password]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"authType={param_list[0]}&userName={param_list[1]}&password={param_list[2]}"
        address = httpLocation + params
    
        result = datahub.get(address, token=self._token)
        return result
    
    # def xmldownload(self, filePathAsBase64: str, writeFullPath:str):
    #     """Download XML file for the current ETAP project.
        
    #     ArgsL
    #        filePathAsBase64 (str):     Encrypted full path to where file should be written to
    #        writeFullPath (str):        Full path to where file should be written to"""
    #     # return datahub.get_file(f'{self._baseAddress}{apiconstants.projectdata_xmldownload}/{filePathAsBase64}', writeFullPath, token=self._token)
    #     httpLocation = f'{self._baseAddress}{apiconstants.projectdata_xmldownload}?fullPathToFileAsBase64='
    #     params = filePathAsBase64
    #     address = httpLocation + urllib.parse.quote(params)
    #     # return datahub.get_file(f'{self._baseAddress}{apiconstants.projectdata_xmldownload}/{filePathAsBase64}', writeFullPath, token=self._token)
    #     return datahub.get_file(address, writeFullPath, token=self._token)
    
    # def xmldownloadflat(self, filePathAsBase64: str, writeFullPath:str):
    #     """Download Flat XML file for the current ETAP project.
        
    #     ArgsL
    #        filePathAsBase64 (str):     Encrypted full path to where file should be written to
    #        writeFullPath (str):        Full path to where file should be written to"""
    #     httpLocation = f'{self._baseAddress}{apiconstants.projectdata_xmldownloadflat}?fullPathToFileAsBase64='
    #     params = filePathAsBase64
    #     address = httpLocation + urllib.parse.quote(params)
    #     return datahub.get_file(address, writeFullPath, token=self._token)

    def activateelement(self, presentation: str, component: str) -> str:

        """select the element or composite in ETAP

        Args:
            presentation (str): Presentation name
            component (str): Component name

        Returns:
            str: 'True' for success or 'False' for failure
        """       

        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_activateelement}/{presentation}/{component}'
    
        emptyDict = dict()
        result = datahub.post(httpLocation, emptyDict, token=self._token)
        return result

    def isprojectauthenticationrequired(self) -> str:
        """Returns TRUE if project authentication is required or FALSE if not required"""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_isprojectauthenticationrequired}', token=self._token)

    def setelementsprops(self, compositeNetwork, MultipleElementJson=None ) -> str:
        """Set multiple field of multiple elements at once

        Args:
            compositeNetwork (str): Composite Network name for element properties to be updated.
            MultipleElementJson (json): Json input to update multiple field of elements at once.
            Element Type: BUS,CABLE (getelementtypes endpoint can be utilized to get the names of supported element types)
            Element Name: Main Bus, LV Bus (getelementnames endpoint can be utilized to get the names of the elemets)
            Field Name: ID, NominalKV (getelementpropertynamesxml endpoint can be utilized to get the field names)
            NOTE: Elements inside same composite network can only be updated. For elements on Main OLV, keep the parameter blank. 
        Returns:
            str: 'True' for success or 'False' for failure
        """       

        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_setelementsprops}?'
        
        param_list = [compositeNetwork]
        
        # URL encode all the params
        for elem in param_list:
            param_list[param_list.index(elem)] = urllib.parse.quote(elem)
        
        params = f"compositeNetwork={param_list[0]}"
        address = httpLocation + params
        result = datahub.post(address, MultipleElementJson, token=self._token)
        return result

    def getprotectedelements(self, relayName:str) -> str:
        """Returns the list of protected elements given a relay name"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getprotectedelements}?relayName='
        params = relayName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def getreferencebus(self, relayName:str) -> str:
        """Returns reference bus given a relay name"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getreferencebus}?relayName='
        params = relayName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def gettiebuses(self, busName:str) -> str:
        """Returns the list of tie buses given a bus name"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_gettiebuses}?busName='
        params = busName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def iselementtransformer(self, elementName:str) -> str:
        """Returns TRUE if element is a transformer or FALSE if element is not a transformer"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementtransformer}?elementName='
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def iselementdoubleline(self, elementName:str) -> str:
        """Returns TRUE if element is a doubleline or FALSE if element is not a doubleline"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_iselementdoubleline}?elementName='
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def getshortestremotelineimp(self, branchName:str, refBusName:str) -> str:
        """Returns shortest remote line impedance"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getshortestremotelineimp}'
       
        params = f"?branchName={branchName}&refBusName={refBusName}"
        address = httpLocation + params
        return datahub.get(address, token=self._token)
    
    def getlongestremotelineimp(self, branchName:str, refBusName:str) -> str:
        """Returns longest remote line impedance"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getlongestremotelineimp}'
       
        params = f"?branchName={branchName}&refBusName={refBusName}"
        address = httpLocation + params
        return datahub.get(address, token=self._token)
    
    def getbranchimpedance(self, branchName:str, refBusName:str, remoteBusName:str, seriesCapFlag=None) -> str:
        """Returns the list of branch impedance"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getbranchimpedance}'
       
        params = f"?branchName={branchName}&refBusName={refBusName}&remoteBusName={remoteBusName}&seriesCapFlag={seriesCapFlag}"
        address = httpLocation + params
        return datahub.get(address, token=self._token)
    
    def getratedkv(self, elementName:str) -> str:
        """Returns the following attached bus kV information for the specified element ID.

            primaryKv - From bus kV for branch elements and attached bus kV for non-branch elements.

            secondaryKv – To bus kV for branch elements and empty string for non-branch elements.

            teritiaryKv – Tertiary bus kV for 3-winding transformers and empty string for other elements.

            Note: Branch elements corresponds lines, cables, transformers"""
            
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getratedkv}?elementName='
        params = elementName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def getadjacentelements(self, branchName:str, remoteBusName:str) -> str:
        """Returns the list of adjacent elements"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getadjacentelements}'
       
        params = f"?branchName={branchName}&remoteBusName={remoteBusName}"
        address = httpLocation + params
        return datahub.get(address, token=self._token)

    def getbusinfo(self, branchName:str) -> str:
        """Returns the following attached bus ID for the specified element ID

        fromBus - From bus for branch elements and attached bus for non-branch elements.

        toBus – To bus for branch elements and empty string for non-branch elements.

        teritiary – Tertiary bus for 3-winding transformers and empty string for other elements.

        Note: Branch elements corresponds lines, cables, transformers"""
            
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getbusinfo}?branchName='
        params = branchName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def getrelaycts(self, relayName:str) -> str:
        """Returns relay CTs"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getrelaycts}?relayName='
        params = relayName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def getrelaypts(self, relayName:str) -> str:
        """Returns relay PTs"""
        httpLocation = f'{self._baseAddress}{apiconstants.projectdata_getrelaypts}?relayName='
        params = relayName
        address = httpLocation + urllib.parse.quote(params)
        return datahub.get(address, token=self._token)
    
    def getrtsleaccuracy(self) -> str:
        """Returns the ETAP real-time accuracy of the single-line editor (e.g., alerts, alarms, warnings, etc.)"""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getrtsleaccuracy}', token=self._token)
    
    
    def getrtstatus(self) -> str:
        """Returns the ETAP real-time status (online or offline)"""
        return datahub.get(f'{self._baseAddress}{apiconstants.projectdata_getrtsleaccuracy}', token=self._token)
