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

from etap.python_engine import mappings
from typing import Mapping
from etap.python_engine.etapelementsclass import Branch, Breaker, Bus, Init_Data, LumpedLoad, Network, Runtime_Data, Utility
from etap.python_engine import pythonengine_pb2
from etap.python_engine import pythonengine_pb2_grpc
import grpc
from concurrent import futures
import xml.dom.minidom as md
import datetime
import sys


class PythonEngineServiceServicer(pythonengine_pb2_grpc.PythonEngineServiceServicer):
    
    def __init__(self) -> None:
        super().__init__()
        self._init_data = Init_Data()
        self._nNetwork = Network()
        
    def PingService(self, request, context):
        resp = pythonengine_pb2.PingResult()
        resp.result = True
        return resp
    
    def Init(self, request, context):
        self._update_initdata(request)
        sys.path.append(user_folder_path)
        from user_code import Init 
        Init(self._init_data)
        
        config_file_path = user_folder_path + "\PythonEngine.config"
        runtime_data = read_config_file(config_file_path)
        self._nNetwork = runtime_data.Network
        
        resp = pythonengine_pb2.InitResult()
        mappings.map_Network_to_gNetwork(u=self._nNetwork, y=resp.Network)
        return resp
    
    def RunStep(self, request, context):
        # gNetwork --> nNetwork
        gNetwork = request.Network
        nNetwork = self._nNetwork  # retrieve from stash
        mappings.map_gNetwork_to_Network(u=gNetwork, y=nNetwork)
        
        # call user code
        r = Runtime_Data()
        r.SimTime = request.SimTime
        r.Network = nNetwork
        sys.path.append(user_folder_path)
        from user_code import RunStep 
        runtime_data = RunStep(r)
        nNetwork = runtime_data.Network
        self._nNetwork = nNetwork  # stash
        
        # gNetwork <-- nNetwork
        resp = pythonengine_pb2.RunStepResult()
        mappings.map_Network_to_gNetwork(u=nNetwork, y=resp.Network)
        return resp
    
    
    def _update_network(self, runstep_request: pythonengine_pb2.RunStepRequest):
        gNetwork = runstep_request.Network
        pNetwork = self._nNetwork
        mappings.map_gNetwork_to_Network(u=gNetwork, y=pNetwork)
    
    
    def _update_initdata(self, init_request: pythonengine_pb2.InitRequest):
        self._init_data.StudyType = init_request.StudyType
        self._init_data.TimeStep = init_request.TimeStep
        self._init_data.TStop = init_request.TStop
        self._init_data.NominalFreq = init_request.NominalFreq       
        
def start_engine(port:int, user_folder):
    global user_folder_path 
    user_folder_path = user_folder
    server = grpc.server(futures.ThreadPoolExecutor(max_workers=10))
    pythonengine_pb2_grpc.add_PythonEngineServiceServicer_to_server(
        PythonEngineServiceServicer(), server)
    server.add_insecure_port(f'[::]:{port}')
    server.start()
    server.wait_for_termination()

def read_config_file(xmlfile):
    file = md.parse(xmlfile)
    init_data = create_init_data_from_xml(file)
    runtime_data = create_runtime_data_from_xml(file)
    return runtime_data

def create_init_data_from_xml(file):
    Init_xml_tags = file.getElementsByTagName("PythonEngine")
    StudyType = [x.attributes["StudyType"].nodeValue for x in Init_xml_tags]
    Freq = [x.attributes["NominalFreq_Hz"].nodeValue for x in Init_xml_tags]
    TimeStep = [x.attributes["TimeStepMs"].nodeValue for x in Init_xml_tags]
    TStop = [x.attributes["TstopSecs"].nodeValue for x in Init_xml_tags]
    init_Data  = Init_Data()
    init_Data.StudyType = StudyType[0]
    init_Data.NominalFreq =Freq[0]
    init_Data.TimeStep = TimeStep[0]
    init_Data.TStop = TStop[0]
    
    return init_Data

def create_runtime_data_from_xml(file):
    
    Runtime_data = Runtime_Data()
    
    #Append Busses
    Get_Busses = file.getElementsByTagName("Buses")
    Bus_names = [x.attributes["IDs"].nodeValue for x in Get_Busses][0].split(',')
    for item in Bus_names:
        Bus_class = Bus()
        Bus_class.In_ID = item
        Runtime_data.Network.Buses.append(Bus_class)
        
    #Append Breakers
    Get_Breakers = file.getElementsByTagName("Breakers")
    Breaker_names = [x.attributes["IDs"].nodeValue for x in Get_Breakers][0].split(',')
    for item in Breaker_names:
        Breaker_class = Breaker()
        Breaker_class.In_ID = item
        Runtime_data.Network.Breakers.append(Breaker_class)
        
    #Append Utilities
    Get_Utilities = file.getElementsByTagName("Utilities")
    Utility_names = [x.attributes["IDs"].nodeValue for x in Get_Utilities][0].split(',')
    for item in Utility_names:
        Utility_class = Utility()
        Utility_class.In_ID = item
        Runtime_data.Network.Utilities.append(Utility_class)
        
    #Append Branches
    Get_Branches = file.getElementsByTagName("Branches")
    Branch_names = [x.attributes["IDs"].nodeValue for x in Get_Branches][0].split(',')
    for item in Branch_names:
        Branch_class = Branch()
        Branch_class.In_ID = item
        Runtime_data.Network.Branches.append(Branch_class)
        
    #Append Lump loads
    Get_LumpedLoads = file.getElementsByTagName("LumpedLoads")
    LumpedLoad_names = [x.attributes["IDs"].nodeValue for x in Get_LumpedLoads][0].split(',')
    for item in LumpedLoad_names:
        LumpedLoad_class = LumpedLoad()
        LumpedLoad_class.In_ID = item
        Runtime_data.Network.LumpedLoads.append(LumpedLoad_class)
    
    return Runtime_data

# if __name__ == "__main__":
#     port=50051
#     user_folder = r"D:\ETAPS\cypress5\Example-Other\Example-PythonEngine\python_user_folder"
#     python_banner = "ETAP Python Engine"
#     print(python_banner)
#     print(f"PE> Started at {datetime.datetime.now()}...")
#     print(f"PE> Listening on port {port}...")
#     # start_engine(port, user_folder)
    