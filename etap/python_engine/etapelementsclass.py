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


from dataclasses import dataclass
from typing import List
from dataclasses import field
import random
import string

class ComplexValues:
    def __init__(self) -> None:
        self.Imaginary = 0.0
        self.Magnitude = 0.0
        self.PhaseDegs = 0.0
        self.Real = 0.0

class ThreePhase():
    def __init__(self) -> None:
        self.A = ComplexValues()
        self.AB = ComplexValues()
        self.B = ComplexValues()
        self.BC = ComplexValues()
        self.C = ComplexValues()
        self.CA = ComplexValues()
        self.Neg = ComplexValues()
        self.Pos = ComplexValues()
        self.Zero = ComplexValues()

def append_random_string(element_name):
    letters = string.ascii_letters
    updated_name = element_name.join(random.choice(letters) for i in range(5))
    return updated_name

class Utility:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Utility")
        self.Energized = True
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_Freq =0.0
        self.In_S_MVA = ComplexValues()
        self.In_S3Phase_MVA = ThreePhase()
             
class Bus:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Bus")
        self.Energized = True
        self.In_V_LG_kV = ThreePhase()
        self.In_Freq = 0.0

class Breaker:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Breaker")
        self.Energized = True
        self.In_status = "Close"
        self.In_Freq = 0.0
        self.Out_status = "Close"

class Branch:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Line")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW_from = 0.0
        self.In_MW_to = 0.0
        self.In_Mvar_from = 0.0
        self.In_Mvar_to = 0.0
        self.In_V_LG_kV_from = ThreePhase()
        self.In_V_LG_kV_to = ThreePhase()   
        self.In_I_kA_from = ThreePhase()
        self.In_I_kA_to = ThreePhase()   
        self.In_S_MVA_from = ComplexValues()
        self.In_S_MVA_to = ComplexValues() 
    
class LumpedLoad:
    def __init__(self) -> None:
        self.In_ID = append_random_string("LumpedLoad")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_PowerFactor = 0.0
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()

class Transformer_2W:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Xfmr_2W")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW_Primary = 0.0
        self.In_MW_Secondary = 0.0
        self.In_Mvar_Primary = 0.0
        self.In_Mvar_Secondary = 0.0
        self.In_PrimaryTap = 0.0
        self.In_SecondaryTap = 0.0
        self.Out_PrimaryTap = 0.0
        self.Out_SecondaryTap = 0.0
        self.In_I_kA_Primary = ThreePhase()
        self.In_I_kA_Secondary = ThreePhase()
        self.In_V_LG_kV_Primary = ThreePhase()
        self.In_V_LG_kV_Secondary = ThreePhase()
        self.In_S_MVA_Primary = ComplexValues()
        self.In_S_MVA_Secondary = ComplexValues() 


class Transformer_3W:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Xfmr_3W")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW_Primary = 0.0
        self.In_MW_Secondary = 0.0
        self.In_MW_Tertiary = 0.0
        self.In_Mvar_Primary = 0.0
        self.In_Mvar_Secondary = 0.0
        self.In_Mvar_Tertiary = 0.0
        self.In_PrimaryTap = 0.0
        self.In_SecondaryTap = 0.0
        self.In_TertiaryTap = 0.0
        self.In_I_kA_Primary = ThreePhase()
        self.In_I_kA_Secondary = ThreePhase()
        self.In_I_kA_Tertiary = ThreePhase()
        self.In_V_LG_kV_Primary = ThreePhase()
        self.In_V_LG_kV_Secondary = ThreePhase()
        self.In_V_LG_kV_Tertiary = ThreePhase()
        self.In_S_MVA_Primary = ComplexValues()
        self.In_S_MVA_Secondary = ComplexValues() 
        self.In_S_MVA_Tertiary = ComplexValues() 

class Capacitor:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Capacitor")
        self.Energized = True
        self.In_Freq = 0.0    
        self.In_PowerFactor = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_Number_of_Banks = 0
        self.Out_Number_of_Banks = 0
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()

class Generator:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Generator") 
        self.Energized = True
        self.In_Freq = 0.0
        self.In_Ifd = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_PowerFactor = 0.0
        self.In_Speed = 0.0
        self.Out_OperationMode = "Swing"
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.Out_PowerFactor = 0.0
        self.Out_Turbine_MW_Max= 0.0
        self.Out_Turbine_V_Pu= 0.0
        self.Out_Turbine_V_AngleDegs= 0.0
        self.Out_Turbine_MW_Max= 0.0
        self.Out_Speed = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()

class DC2DCConverter:
    def __init__(self) -> None:
        self.In_ID  = append_random_string("DC2DCConverter")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_I_kA_Input = 0.0
        self.In_I_kA_Output = 0.0
        self.In_V_LG_kV_Input = 0.0
        self.In_V_LG_kV_Output = 0.0

class DCBattery:
    def __init__(self) -> None:
        self.In_ID = append_random_string("DCBattery")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_I_kA = 0.0
        self.In_V_LG_kV = 0.0   
        self.In_MW = 0.0
        self.In_SOC = 0.0
        self.Out_MW = 0.0

class DCBranch:
    def __init__(self) -> None:
        self.In_ID = append_random_string("DCBattery")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_I_kA_From = 0.0
        self.In_I_kA_To = 0.0
        self.In_V_LG_kV_From = 0.0   
        self.In_V_LG_kV_To = 0.0   

class DCBus:
    def __init__(self) -> None:
        self.In_ID = append_random_string("DCBattery")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_V_LG_kV = 0.0      

class DCLoad:
    def __init__(self) -> None:
        self.In_ID = append_random_string("DCBattery")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_I_kA = 0.0
        self.In_V_LG_kV = 0.0   
        self.In_MW = 0.0

class DCLumpedLoad:
    def __init__(self) -> None:
        self.In_ID = append_random_string("DCBattery")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_I_kA = 0.0
        self.In_V_LG_kV = 0.0   
        self.In_MW = 0.0
        self.Out_MW = 0.0  

class Inverter:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Inverter")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_I_kA_DC = 0.0
        self.In_V_LG_kV_DC = 0.0
        self.In_PowerFactor = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0  
        self.Out_OperationMode = "Swing"
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.Out_PowerFactor = 0.0
        self.Out_Vref_Pu = 0.0
        self.Out_V_AC_AngleDegs = 0.0
        self.Out_I_AC_Zero_kA_Real = 0.0
        self.Out_I_AC_Zero_kA_Imaginary = 0.0
        self.Out_I_AC_Negative_kA_Real = 0.0
        self.Out_I_AC_Negative_kA_Imaginary = 0.0
        self.Out_I_AC_Positive_kA_Real = 0.0
        self.Out_I_AC_Positive_kA_Imaginary = 0.0
        self.In_I_kA = ThreePhase()  
        self.In_V_LG_kV = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()           


class StaticLoad:
    def __init__(self) -> None:
        self.In_ID = append_random_string("StaticLoad")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_PowerFactor = 0.0
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()
        
class Multimeter:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Multimeter")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_PowerFactor = 0.0
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()

class PV:
    def __init__(self) -> None:
        self.In_ID = append_random_string("PV")
        self.Energized = True
        self.In_Freq = 0.0     
        self.In_I_kA_DC = 0.0
        self.In_V_LG_kV_DC = 0.0               
        self.In_Irradiance = 0.0  
        self.In_MW = 0.0
        self.In_MW_Available = 0.0          
        self.In_Temperature_Degrees = 0.0
        self.Out_MW = 0.0

class PVInverter:
    def __init__(self) -> None:
        self.In_ID = append_random_string("PVInverter")
        self.Energized = True
        self.In_Freq = 0.0     
        self.In_I_kA_DC = 0.0
        self.In_V_LG_kV_DC = 0.0               
        self.In_Irradiance = 0.0  
        self.In_MW = 0.0
        self.In_PowerFactor = 0.0
        self.In_Mvar = 0.0
        self.In_MW_Available = 0.0          
        self.In_Temperature_Degrees = 0.0
        self.Out_OperationMode = "Swing"
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.Out_PowerFactor = 0.0
        self.Out_Vref_Pu = 0.0
        self.Out_I_kA = 0.0
        self.Out_V_LG_kV = 0.0
        self.Out_V_AC_AngleDegs = 0.0
        self.Out_I_AC_Zero_kA_Real = 0.0
        self.Out_I_AC_Zero_kA_Imaginary = 0.0
        self.Out_I_AC_Negative_kA_Real = 0.0
        self.Out_I_AC_Negative_kA_Imaginary = 0.0
        self.Out_I_AC_Positive_kA_Real = 0.0
        self.Out_I_AC_Positive_kA_Imaginary = 0.0
        self.Out_Irradiance = 0.0 
        self.In_I_kA = ThreePhase()  
        self.In_V_LG_kV = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()   
        
class Statcom:
    def __init__(self) -> None:
        self.In_ID = append_random_string("Statcom")
        self.Energized = True
        self.In_Freq = 0.0     
        self.In_I_kA_DC = 0.0
        self.In_V_LG_kV_DC = 0.0               
        self.In_MW = 0.0
        self.In_PowerFactor = 0.0
        self.In_Mvar = 0.0         
        self.Out_OperationMode = "Swing"
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.Out_PowerFactor = 0.0
        self.Out_Vref_Pu = 0.0
        self.Out_V_AC_AngleDegs = 0.0
        self.Out_I_AC_Zero_kA_Real = 0.0
        self.Out_I_AC_Zero_kA_Imaginary = 0.0
        self.Out_I_AC_Negative_kA_Real = 0.0
        self.Out_I_AC_Negative_kA_Imaginary = 0.0
        self.Out_I_AC_Positive_kA_Real = 0.0
        self.Out_I_AC_Positive_kA_Imaginary = 0.0
        self.In_I_kA = ThreePhase()  
        self.In_V_LG_kV = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()  

class UPS:
    def __init__(self) -> None:
        self.In_ID = append_random_string("UPS")
        self.Energized = True
        self.In_Freq = 0.0     
        self.In_I_kA_DC = 0.0
        self.In_V_LG_kV_DC = 0.0              
        self.In_MW_Input = 0.0
        self.In_Mvar_Input = 0.0  
        self.In_MW_Output = 0.0
        self.In_Mvar_Output = 0.0          
        self.In_MW_DC = 0.0          
        self.In_I_kA_Input = ThreePhase()  
        self.In_I_kA_Output = ThreePhase()  
        self.In_V_LG_kV_Input = ThreePhase()
        self.In_V_LG_kV_Output = ThreePhase()
        self.In_S_MVA_Input = ComplexValues()   
        self.In_S_MVA_output = ComplexValues()   

class WTG:
    def __init__(self) -> None:
        self.In_ID = append_random_string("WTG")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_PowerFactor = 0.0
        self.In_MW_Available = 0.0
        self.In_WindSpeed = 0.0
        self.Out_OperationMode = "Swing"
        self.Out_MW = 0.0
        self.Out_Mvar = 0.0
        self.Out_PowerFactor = 0.0
        self.Out_Vref_Pu = 0.0
        self.Out_V_AC_AngleDegs = 0.0
        self.Out_I_kA = 0.0
        self.Out_V_LG_kV = 0.0
        self.Out_WindSpeed = 0.0
        self.Out_I_AC_Zero_kA_Real = 0.0
        self.Out_I_AC_Zero_kA_Imaginary = 0.0
        self.Out_I_AC_Negative_kA_Real = 0.0
        self.Out_I_AC_Negative_kA_Imaginary = 0.0
        self.Out_I_AC_Positive_kA_Real = 0.0
        self.Out_I_AC_Positive_kA_Imaginary = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()

class InductionMotor:
    def __init__(self) -> None:
        self.In_ID = append_random_string("InductionMotor")
        self.Energized = True
        self.In_Freq = 0.0
        self.In_MW = 0.0
        self.In_Mvar = 0.0
        self.In_PowerFactor = 0.0
        self.In_Slip = 0.0
        self.Out_LoadingPercent = 0.0
        self.In_V_LG_kV = ThreePhase()
        self.In_I_kA = ThreePhase()
        self.In_S3Phase_MVA = ThreePhase()
        self.In_S_MVA = ComplexValues()
         
class DoubleThrowSwitch:
    def __init__(self) -> None:
        self.In_ID = append_random_string("DTSwitch")
        self.Energized = True
        self.In_status = "Close"
        self.In_Freq = 0.0
        self.Out_status = "Close"
                              
@dataclass
class Network:
    Utilities : List[Utility] = field(default_factory=list)
    Buses : List[Bus] = field(default_factory=list)
    Breakers : List[Breaker]= field(default_factory=list)
    Branches : List[Branch]= field(default_factory=list)
    LumpedLoads : List[LumpedLoad]= field(default_factory=list)
    Transformers_2Ws :List[Transformer_2W] = field(default_factory=list)
    Transformers_3Ws :List[Transformer_3W] = field(default_factory=list)
    Capacitors :List[Capacitor] = field(default_factory=list)
    Generators :List[Generator] = field(default_factory=list)
    DC2DCConverters :List[DC2DCConverter] = field(default_factory=list)
    DCBatteries :List[DCBattery] = field(default_factory=list)
    DCBranches :List[DCBranch] = field(default_factory=list)
    DCBuses :List[DCBus] = field(default_factory=list)
    DCLoads :List[DCLoad] = field(default_factory=list)
    DCLumpedLoads :List[DCLumpedLoad] = field(default_factory=list)
    Inverters :List[Inverter] = field(default_factory=list)
    Loads :List[StaticLoad] = field(default_factory=list)
    Multimeters :List[Multimeter] = field(default_factory=list)
    PVs :List[PV] = field(default_factory=list)
    PVInverters :List[PVInverter] = field(default_factory=list)
    Statcoms :List[Statcom] = field(default_factory=list)
    UPSes :List[UPS] = field(default_factory=list)
    WTGs :List[WTG] = field(default_factory=list)
    InductionMotors :List[InductionMotor] = field(default_factory=list)
    DoubleThrowSwitches :List[DoubleThrowSwitch] = field(default_factory=list)
    
@dataclass
class Init_Data:   
    StudyType:str = "Dummy"
    TimeStep:float = 1
    TStop:float = 10.0
    NominalFreq:float = 0.0

class Runtime_Data:
    def __init__(self) -> None:
        self.SimTime = 0.0
        self.Network = Network()

# @dataclass
# class Runtime_Data:
#    SimTime:float = 0.0
#    Network:Network = Network()

@dataclass  
class PythonEngine:
    Version:str = "1.11"
    Init:Init_Data = Init_Data()
    Runtime:Runtime_Data = Runtime_Data()

