# Author - G.Kidwell.   
#
# Modified from Intel Preemptive_SVID_Example.py
#


import clr
import sys, os
import time

#sys.path.append(r'C:\Users\gkidw\Gen5_python36_32\GEN5\PythonScripts')
sys.path.append(os.path.join(sys.path[0],'intel'))

from Common.APIs import (DataAPI, VectorAPI, MainAPI, DisplayAPI, MeasurementAPI, GeneratorAPI)
from Common.Models import (TriggerModel, FanModel, SVIDModel, RawDataModel, DataModel, DisplayModel)
from Common.Enumerations import (HorizontalScale, ScoplessChannel, Cursors, PowerState, Transition,
                                Protocol, RailTestMode, FRAME, VECTORINDEX, SVIDCMD,
                                DACTRIGCHN, SPARITY, VSTATUS, DACTRIGCHN, SVIDBUSVECTORS, SVIDBURSTTRIGGER)

from System import String, Char, Int32, UInt16, Boolean, Array, Byte, Double
from System.Collections.Generic import List
from System import Enum

import openpyxl

generator_api   = GeneratorAPI()
display_api     = DisplayAPI()
data_api        = DataAPI()
measurement_api = MeasurementAPI()
vector_api      = VectorAPI()
main_api        = MainAPI()

def convert_voltage_to_vid(generator_api, rail_name, voltage):
    v_step = 0.005
    min_voltage = 0.25
    return ((voltage - min_voltage) / v_step)

def ClearAlert():
    data_api.SetSvidCmdWrite(0,7,0x10,1,5,2,1,3)


def setup_tool_display(testrail,startvoltage,testcurrent):
    generator_api.AssignRailToDriveOne(testrail)
    display_api.Ch1_2Rail(testrail)
    display_api.DisplayDigitalSignal("SVID1ALERT")
    display_api.DisplayDigitalSignal("CSO1#")
    display_api.SetChannel(ScoplessChannel.Ch1, True)
    display_api.SetChannel(ScoplessChannel.Ch2, True)
    display_api.SetHorizontalScale(HorizontalScale.Scale20us)
    max_voltage = max(startvoltage,0)
    min_voltage = 0
    display_api.SetVerticalVoltageScale(min_voltage * 0.8, max_voltage * 1.2)
    display_api.SetVerticalCurrentScale(testcurrent * .8, testcurrent * 1.2)
    
def setup_tool_measurements(testrail):
    measurement_api.ClearAllMeasurements();               time.sleep(0.1)
    measurement_api.MeasureCSOAlert(testrail);            time.sleep(0.1)
    measurement_api.MeasureVoltageMinMax(testrail);       time.sleep(0.1)
    measurement_api.MeasureVoltageRiseFallTime(testrail); time.sleep(0.1)
    measurement_api.MeasureCSOAlert(testrail);            time.sleep(0.1)
    measurement_api.MeasureCurrentMean(testrail);         time.sleep(0.1)

def setup_initial_voltage_and_current(testrail,startvoltage,testcurrent):
    generator_api.SetVoltageForRail(testrail, startvoltage, Transition.Fast); time.sleep(0.5)
    enable_current_tracking=True
    generator_api.SetTrackingEnabled(testrail, enable_current_tracking); time.sleep(1)
    generator_api.Generator1SVSC(testrail, testcurrent, True);                time.sleep(3) 
    ClearAlert();                                                             time.sleep(0.5)

def setup_and_run_vectors(testrail):
    vector_api.ResetVectors()  # Deactivates all vectors, run before and after using vectors
    vector_api.VectorTriggerSettingsInternal(0, False)  # Repeat delay, continous repeating
    # Get the rail information
    rails = generator_api.GetRails()
    rail = None
    for r in rails:
        if r.Name == testrail:
            rail = r
    if rail is None:
        print('Test failed because the test rail could not be found. Please check the rail name.')
        exit()


    # CreateSimpleVector(
    # VECTORINDEX Vindex = VECTORINDEX.VECTOR1,     # Vector index
    # ushort vrAddr = 0,                            # VR Address
    # SVIDCMD svidCmd = SVIDCMD.FAST,               # SVID Command
    # byte data= 0x30,                              # Payload
    # ushort delay = 2                              # VR clock cycle delay before next vector
    # )
    vector_api.CreateSimpleVector(1, rail.VRAddress,SVIDCMD.GETREG,0x31,20 ) # Read start VID
    vector_api.CreateSimpleVector(2, rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read start power state
    vector_api.CreateSimpleVector(3, rail.VRAddress,SVIDCMD.FAST,  0x00,800) # SetVIDFast to 00h  - NOTE Alert# should remain high
    vector_api.CreateSimpleVector(4, rail.VRAddress,SVIDCMD.GETREG,0x10,20 ) # get Status1 - should be VR_Settled
    vector_api.CreateSimpleVector(5, rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read end power state register
    vector_api.CreateSimpleVector(6, rail.VRAddress,SVIDCMD.GETREG,0x31,20 ) # Read end VID
    vector_api.CreateSimpleVector(7, rail.VRAddress,SVIDCMD.PS,    0x03,20 ) # Set PS to 3 - should ACK
    vector_api.CreateSimpleVector(8, rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read PS - should read PS0
    vector_api.CreateSimpleVector(9, rail.VRAddress,SVIDCMD.FAST,  0x82,800) # SetVIDfast to 900mV
    vector_api.CreateSimpleVector(10,rail.VRAddress,SVIDCMD.GETREG,0x10,20 ) # get Status1 - should be VR_Settled
    vector_api.CreateSimpleVector(11,rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read PS - should read PS0
    vector_api.CreateSimpleVector(12,rail.VRAddress,SVIDCMD.GETREG,0x31,20 ) # Read end VID

    vector_api.LoadAllVectors(0, rail.SVIDBus)  # Number of repeats, svid bus
    data_api.SetTrigger(12, 100, 0, 0, 0, False)  # set to svid burst trigger
    time.sleep(1)
    vector_api.ExecuteVectors(SVIDBURSTTRIGGER.INTERNAL)  # Execute the vectors
    vector_api.ReadVectorResponses()  # Update the response data for the vectors
    return data_api.GetVectorDataOnce()  # Get the data for the vector responses
 

def output_tool_display():
    pass

def output_screen(vresult):
    print("------------------------------------------------------------------------------")
    print("                              Hawkmock Vector                       ")
    print("------------------------------------------------------------------------------")
    print("")
    # Prints the SVID responses to vector 1 and 2
    print('ACK\tDATA\tPRTYERR\tPRTYVAL')
    print('-----------------------------')
    for i in range(0, 13):
        print(str(vresult.AckResp[i]),
              '\t', str(vresult.DataResp[i]),
              '\t', str(vresult.ParityErrResp[i]),
              '\t', str(vresult.ParityValResp[i]))

def output_excel_file():
    pass

def shutdown_tool(testrail):
    vector_api.ResetVectors()
    # Shutdown the generator
    #generator_api.Generator1SVSC(test_rail, test_current, False)
    generator_api.Generator1SVSC(testrail,0,False)
    generator_api.SetTrackingEnabled(testrail, False)

    # Erase the vector data from the GUI and return to normal operation
    #data_api.SetTrigger(4, 0, 0, 0, 0, False)  # GUI will stay on the vector data if you comment this out




#-------------------------------------------------------------------------
# Input variables 

test_rail     = 'VCCCORE'
start_voltage = 0.9
test_current  = 20  # (A)

#-------------------------------------------------------------------------
# Main program

setup_tool_display(test_rail,start_voltage,test_current)
#setup_tool_measurements()
setup_initial_voltage_and_current(test_rail,start_voltage,test_current)
vResult=setup_and_run_vectors(test_rail)
output_screen(vResult)
shutdown_tool(test_rail)

#
#--------------------------------------------------------------------------






