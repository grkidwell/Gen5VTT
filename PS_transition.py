# ---------------------------- GEN5 VRTT - PS Transition Example ---------------------
# ------------------------------------------------------------------------------------
#   Intel Confidential 
# ------------------------------------------------------------------------------------
#
# - Adjust the initial_ps and next_ps variables to try different PS states
# ------------------------------------------------------------------------------------
#   Rev 1.0   - 5/15/2018
#           - Initial versioning.
#   Rev 1.1 - 11/7/2019
#           - Updated and cleaned
#   pulled latest from GEN5CONTROLLER\NEXTGEN\PythonApplication2

import clr
import sys,os
import time

sys.path.append(os.path.join(sys.path[0],'intel'))


from Common.APIs import DataAPI
from Common.APIs import VectorAPI
from Common.APIs import MainAPI
from Common.APIs import DisplayAPI
from Common.Models import TriggerModel
from Common.Models import FanModel
from Common.Models import SVIDModel
from Common.Models import RawDataModel
from Common.Models import DataModel
from Common.Models import DisplayModel
from Common.Enumerations import HorizontalScale
from Common.Enumerations import ScoplessChannel
from Common.Enumerations import Cursors
from Common.APIs import MeasurementAPI
from Common.APIs import GeneratorAPI
#from CommonBL import *
from Common.Enumerations import PowerState
from Common.Enumerations import Transition
from Common.Enumerations import Protocol
from Common.Enumerations import RailTestMode
from Common.Enumerations import FRAME
from Common.Enumerations import SVIDBURSTTRIGGER
from Common.Enumerations import ScoplessChannel
from Common.Enumerations import VECTORINDEX
from Common.Enumerations import SVIDCMD
from Common.Enumerations import DACTRIGCHN

from Common.Enumerations import VECTORINDEX 
from Common.Enumerations import FRAME 
from Common.Enumerations import SVIDCMD 
from Common.Enumerations import SPARITY
from Common.Enumerations import VSTATUS
from Common.Enumerations import DACTRIGCHN
from Common.Enumerations import SVIDBUSVECTORS
from Common.Enumerations import SVIDBURSTTRIGGER

#from CommonBL import *
from System import String, Char, Int32, UInt16, Boolean, Array, Byte, Double
from System.Collections.Generic import List

gen = GeneratorAPI()
display = DisplayAPI()
data = DataAPI()
meas = MeasurementAPI()
vector = VectorAPI()
api = MainAPI()

########## End of imports
print("##############################################################################")
print("------------------------------------------------------------------------------")
print("                              Vector PS Transition                            ")
print("------------------------------------------------------------------------------")
print("##############################################################################")
print("")

# Change these as needed
initial_ps =0
next_ps =1
vraddress = 0
svid_bus = 0
test_rail = 'VCCCORE'
test_voltage = 1.2
clock_delay = 2

## What Rails are available?
print ("Displaying rails available")
print ("------------------------------------------"); 
t=gen.GetRails()
rail = None
for i in range (0, len(t)):
    print(t[i].Name, "Vid Table - ", t[i].VID, " Max Current - ", t[i].MaxCurrent, "; Vsense - ", t[i].VSenseName)
    print("Load sections: " + t[i].LoadSections)
    if t[i].Name == test_rail:
        rail = t[i]
        break
print ("")
print ("Begining Test")
print ("--------------")

#####Assign rails to Tabs
gen.AssignRailToDriveOne(test_rail)

# Set voltage and current
#if rail != None:
    #gen.SetVoltageForRail(test_rail, test_voltage, Transition.Fast)
    #gen.Generator1DVSC_I(test_rail, 1, True)






# Display channel
display.Ch1_2Rail(test_rail)

# Display digital signals
display.DisplayDigitalSignal("SVID1ALERT")
display.DisplayDigitalSignal("CSO1#")

# Measurement setup
meas.MeasureVoltageMinMax(test_rail)
V_mean= meas.MeasureVoltageMean(test_rail)

####Enable Scope Traces
display.SetChannel(ScoplessChannel.Ch1, True)
display.SetChannel(ScoplessChannel.Ch2, True)



# public enum SVIDCMD
#     {
#         [Description("EXTENDED")]
#         EXTENDED = 0,
#         [Description("FAST")]
#         FAST = 1,
#         [Description("SLOW")]
#         SLOW = 2,
#         [Description("DECAY")]
#         DECAY = 3,
#         [Description("PS")]
#         PS = 4,
#         [Description("REGADDR")]
#         REGADDR = 5,
#         [Description("REGDATA")]
#         REGDATA = 6,
#         [Description("GETREG")]
#         GETREG = 7,
#         [Description("TESTMODE")]
#         TESTMODE = 8,
#         [Description("SETWP")]
#         SETWP = 9
#     }

vector.ResetVectors()

# CreateSimpleVector(VECTORINDEX Vindex = VECTORINDEX.VECTOR1, ushort vrAddr = 0, SVIDCMD svidCmd = SVIDCMD.FAST,
#            byte data = 0x30, ushort delay = 2)


vector.CreateSimpleVector(1, vraddress, SVIDCMD.PS, initial_ps, 50) # Get status reg
vector.CreateSimpleVector(2, vraddress, SVIDCMD.GETREG, 0x32, 2) # Get status reg
vector.CreateSimpleVector(3, vraddress, SVIDCMD.PS, initial_ps, 2) # Get status reg
vector.CreateSimpleVector(4, vraddress, SVIDCMD.PS, next_ps, 2) # Get status reg
vector.CreateSimpleVector(5, vraddress, SVIDCMD.PS, initial_ps, 2) # Get status reg
vector.CreateSimpleVector(6, vraddress, SVIDCMD.PS, next_ps, 2) # Get status reg
vector.CreateSimpleVector(7, vraddress, SVIDCMD.PS, initial_ps, 2) # Get status reg
vector.CreateSimpleVector(8, vraddress, SVIDCMD.PS, next_ps, 2) # Get status reg
vector.CreateSimpleVector(9, vraddress, SVIDCMD.GETREG, 0x32, 2) # Get status reg




vector.VectorTriggerSettingsMode4(DACTRIGCHN.DAC1, False)
vector.VectorTriggerSettingsInternal(0, False) # Repeat delay, continously repeating
vector.LoadAllVectors(0, svid_bus) # Number of repeats, svid bus
data.SetTrigger(12, 100, 0, 0, 0, False) # set to svid burst trigger

time.sleep(1)

vector.ExecuteVectors(SVIDBURSTTRIGGER.INTERNAL)
vector.ReadVectorResponses()
vector.ResetVectors()

vResult = data.GetVectorDataOnce()
print('ACK\tDATA\tPRTYERR\tPRTYVAL')
print('-----------------------------')
for i in range(0, 9):
    print(str(vResult.AckResp[i]), 
    '\t', str(vResult.DataResp[i]),
        '\t',str(vResult.ParityErrResp[i]),
            '\t',str(vResult.ParityValResp[i]))


print('initial PS = {}'.format(vResult.DataResp[1]))
print('final PS = {}'.format(vResult.DataResp[8]))
gen.Generator1DVSC_I(test_rail, 1, False)

