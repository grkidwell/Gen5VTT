# ---------------------------- GEN5 VRTT - PS Transition Example ---------
# ------------------------------------------------------------------------------------
#   Intel Confidential
# ------------------------------------------------------------------------------------
# - This script sets a start voltage and static current
# - Then the GEN5 issues the "end command" which is preempted by the "preemptive command"
# - Adjust the end command and preemptive command variables to try different PS states or VIDs
# ------------------------------------------------------------------------------------
#   Rev 1.0   - 5/15/2018
#           - Initial versioning.
#   Rev 1.1 - 11/7/2019
#           - Updated and cleaned
#   Rev 1.2 - 4/13/2020
#   Rev 1.3 - 9/28/2020
#           - Removed CSO to alert calculations
#           - Now show the 4th CSO(Preemptive command) to last alert time only
#           - Added a voltage measurement 4us after the last alert
#           - Added current tracking option

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




def add_chart_timing(data: list, time_per_point: float):
    return [
        {'x': idx * time_per_point, 'y': v}
        for idx, v
        in enumerate(data)
    ]

def optimize_chart(points: list):
    new_data= []

    point = None
    value = None
    value_last = None
    for point_next in points:
        value_next = point_next['y']
        if value_last != value or value != value_next:
            if point is not None:
                new_data.append(point)
        value_last = value
        value = value_next
        point = point_next
    new_data.append(point)

    return new_data

def convert_voltage_to_vid(generator_api, rail_name, voltage):
    v_step = 0.005
    min_voltage = 0.25
    return ((voltage - min_voltage) / v_step)

def ClearAlert():
    data_api.SetSvidCmdWrite(0,7,0x10,1,5,2,1,3)

generator_api = GeneratorAPI()
display_api = DisplayAPI()
data_api = DataAPI()
measurement_api = MeasurementAPI()
vector_api = VectorAPI()
main_api = MainAPI()



# End of imports
print("##############################################################################")
print("------------------------------------------------------------------------------")
print("                              Hawkmock Vector                       ")
print("------------------------------------------------------------------------------")
print("##############################################################################")
print("")


# Test input variables - Edit below
#####################################################################
#####################################################################
#####################################################################
test_rail = 'VCCCORE'
start_voltage = 0.9

# Use this for voltage
# command = SVIDCMD.FAST  # SLOW, DECAY, FAST
# command_data = convert_voltage_to_vid(generator_api, test_rail, 1.83)  # Use this function to input by voltage
# OR
# command_data = 0x64  # Input directly if you know the vid code

# Use this for power state transitions
# command = SVIDCMD.PS
# command_data = 1 # 0, 1, 2, 3
end_voltage = 0.8
end_command = SVIDCMD.FAST
end_command_data = convert_voltage_to_vid(generator_api, test_rail, end_voltage)  # Use this function to input by voltage

test_current = 20  # (A)
horizontal_scale = HorizontalScale.Scale20us  # Set the display scale

enable_current_tracking = True # Enables tracking that will auto adjust current level to match current setpoint
#####################################################################
#####################################################################
#####################################################################
# End test input variables - Stop editing


# Get the rail information
rails = generator_api.GetRails()
rail = None
for r in rails:
    if r.Name == test_rail:
        rail = r
if rail is None:
    print('Test failed because the test rail could not be found. Please check the rail name.')
    exit()

# Assign rail to drive tab
generator_api.AssignRailToDriveOne(test_rail)

# Display channel
display_api.Ch1_2Rail(test_rail)
# Display digital signals
display_api.DisplayDigitalSignal("SVID1ALERT")
display_api.DisplayDigitalSignal("CSO1#")
# Enable Scope Traces
display_api.SetChannel(ScoplessChannel.Ch1, True)
display_api.SetChannel(ScoplessChannel.Ch2, True)
# Set horizontal scale
display_api.SetHorizontalScale(horizontal_scale)
# Set vertical scales
max_voltage = max(start_voltage,0)
min_voltage = min(start_voltage, end_voltage, 0)
display_api.SetVerticalVoltageScale(min_voltage * 0.8, max_voltage * 1.2)
display_api.SetVerticalCurrentScale(test_current * .8, test_current * 1.2)

# Measurement setup
measurement_api.ClearAllMeasurements()
time.sleep(0.1)
measurement_api.MeasureCSOAlert(test_rail)
time.sleep(0.1)
measurement_api.MeasureVoltageMinMax(test_rail)
time.sleep(0.1)
measurement_api.MeasureVoltageRiseFallTime(test_rail)
time.sleep(0.1)
measurement_api.MeasureCSOAlert(test_rail)
time.sleep(0.1)
measurement_api.MeasureCurrentMean(test_rail)
time.sleep(0.1)

#Set Initial Voltage and Current
generator_api.SetVoltageForRail(test_rail, start_voltage, Transition.Fast)  # Set start voltage
time.sleep(0.5)
generator_api.SetTrackingEnabled(test_rail, enable_current_tracking)
time.sleep(1)
generator_api.Generator1SVSC(test_rail, test_current, True)  # Set the static current
time.sleep(3)  # Wait a second so that the start voltage is settled
ClearAlert()
time.sleep(0.5)

rail_voltage_current = data_api.GetVoltageCurrent(test_rail, -1)  # Read start voltage
start_voltage_wave = list(rail_voltage_current.Voltage)
mean_start_voltage = sum(
    list(rail_voltage_current.Voltage)) / len(list(rail_voltage_current.Voltage))  # Calculate mean

data_api.SetTrigger(0xC, 1000, 0, 0, 0, False)  # Set the scope trigger to svid burst, offset by 1000 data points for visibility

vector_api.ResetVectors()  # Deactivates all vectors, run before and after using vectors
vector_api.VectorTriggerSettingsInternal(0, False)  # Repeat delay, continous repeating

# CreateSimpleVector(
# VECTORINDEX Vindex = VECTORINDEX.VECTOR1,     # Vector index
# ushort vrAddr = 0,                            # VR Address
# SVIDCMD svidCmd = SVIDCMD.FAST,               # SVID Command
# byte data= 0x30,                              # Payload
# ushort delay = 2                              # VR clock cycle delay before next vector
# )


# Read start VID
vector_api.CreateSimpleVector(
    1,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x31,
    20) 

# Read start power state
vector_api.CreateSimpleVector(
    2,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x32,
    20) 

# SetVIDFast to 00h  - NOTE Alert# should remain high
vector_api.CreateSimpleVector(
    3,
    rail.VRAddress,
    SVIDCMD.FAST,
    0x00,
    800)  

# get Status1 - should be VR_Settled
vector_api.CreateSimpleVector(
    4,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x10,
    20) 

# Read end power state register
vector_api.CreateSimpleVector(
    5,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x32,
    20) 

# Read end VID
vector_api.CreateSimpleVector(
    6,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x31,
    20)

# Set PS to 3 - should ACK
vector_api.CreateSimpleVector(
    7,
    rail.VRAddress,
    SVIDCMD.PS,
    0x03,
    20)

# Read PS - should read PS0
vector_api.CreateSimpleVector(
    8,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x32,
    20)

# SetVIDfast to 900mV
vector_api.CreateSimpleVector(
    9,
    rail.VRAddress,
    SVIDCMD.FAST,
    0x82,
    800)

# get Status1 - should be VR_Settled
vector_api.CreateSimpleVector(
    10,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x10,
    20) 

# Read PS - should read PS0
vector_api.CreateSimpleVector(
    11,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x32,
    20)

# Read end VID
vector_api.CreateSimpleVector(
    12,
    rail.VRAddress,
    SVIDCMD.GETREG,
    0x31,
    20)

vector_api.LoadAllVectors(0, rail.SVIDBus)  # Number of repeats, svid bus
data_api.SetTrigger(12, 100, 0, 0, 0, False)  # set to svid burst trigger

time.sleep(1)

vector_api.ExecuteVectors(SVIDBURSTTRIGGER.INTERNAL)  # Execute the vectors

vector_api.ReadVectorResponses()  # Update the response data for the vectors
vResult = data_api.GetVectorDataOnce()  # Get the data for the vector responses
vector_responses = vResult
# Prints the SVID responses to vector 1 and 2
print('ACK\tDATA\tPRTYERR\tPRTYVAL')
print('-----------------------------')
for i in range(0, 13):
    print(str(vResult.AckResp[i]),
          '\t', str(vResult.DataResp[i]),
          '\t', str(vResult.ParityErrResp[i]),
          '\t', str(vResult.ParityValResp[i]))

vector_api.ResetVectors()

# Shutdown the generator
#generator_api.Generator1SVSC(test_rail, test_current, False)
generator_api.Generator1SVSC(test_rail,0,False)
generator_api.SetTrackingEnabled(test_rail, False)

# Erase the vector data from the GUI and return to normal operation
#data_api.SetTrigger(4, 0, 0, 0, 0, False)  # GUI will stay on the vector data if you comment this out
