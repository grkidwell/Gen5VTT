# Author - G.Kidwell.   
#
# 12/2/20 - Modified from Intel Preemptive_SVID_Example.py
#

import clr,sys,os,time,openpyxl

#sys.path.append(r'C:\Users\gkidw\Gen5_python36_32\GEN5\PythonScripts')
sys.path.append(os.path.join(sys.path[0],'intel'))

from Common.APIs import DataAPI, VectorAPI, MainAPI, DisplayAPI, MeasurementAPI, GeneratorAPI
from Common.Enumerations import HorizontalScale, ScoplessChannel, Transition, SVIDCMD, SVIDBURSTTRIGGER

generator_api   = GeneratorAPI(); display_api     = DisplayAPI()
data_api        = DataAPI();      measurement_api = MeasurementAPI()
vector_api      = VectorAPI();    main_api        = MainAPI()

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
    
def setup_initial_voltage_and_current(testrail,startvoltage,testcurrent):
    generator_api.SetVoltageForRail(testrail, startvoltage, Transition.Fast); time.sleep(0.5)
    enable_current_tracking=True
    generator_api.SetTrackingEnabled(testrail, enable_current_tracking);      time.sleep(1)
    generator_api.Generator1SVSC(testrail, testcurrent, True);                time.sleep(3) 
    data_api.SetSvidCmdWrite(0,7,0x10,1,5,2,1,3); #Clear Alert

def setup_and_run_vectors(testrail,vectordict):
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

    for idx,value in vectordict.items():
        vector_api.CreateSimpleVector(int(idx),rail.VRAddress,value['cmd'],value['data'],value['delay'])

    vector_api.LoadAllVectors(0, rail.SVIDBus)  # Number of repeats, svid bus
    data_api.SetTrigger(12, 100, 0, 0, 0, False)  # set to svid burst trigger
    time.sleep(1)
    vector_api.ExecuteVectors(SVIDBURSTTRIGGER.INTERNAL)  
    vector_api.ReadVectorResponses()  
    return data_api.GetVectorDataOnce()  
 
def output_tool_display():
    pass

def output_screen(vresult,vectordict):
    print("------------------------------------------------------------------------------")
    print("                              Hawkmock Vector                                 ")
    print("------------------------------------------------------------------------------")
    print("")
    print('COMMAND   \tACK\tDATA\tPRTYERR\tPRTYVAL')
    print('------------------------------------------------')
    for i in range(0, len(vectordict.keys())):
        print(vectordict[str(i+1)]['descr'],
              '\t', str(vresult.AckResp[i]),
              '\t', str(vresult.DataResp[i]),
              '\t', str(vresult.ParityErrResp[i]),
              '\t', str(vresult.ParityValResp[i]))

def output_excel_file():
    pass

def shutdown_tool(testrail):
    vector_api.ResetVectors()
    generator_api.Generator1SVSC(testrail,0,False)
    generator_api.SetTrackingEnabled(testrail, False)

    # Erase the vector data from the GUI and return to normal operation
    #data_api.SetTrigger(4, 0, 0, 0, 0, False)  # GUI will stay on the vector data if you comment this out

#-------------------------------------------------------------------------
# Input variables 

test_rail     = 'VCCCORE'
start_voltage = 0.9
test_current  = 20  # (A)

vector_dict = {'1' :{'descr':'read VID  ','cmd':SVIDCMD.GETREG,'data':0x31,'delay':20 },'2' :{'descr':'read PS     ','cmd':SVIDCMD.GETREG,'data':0x32,'delay':20},
               '3' :{'descr':'setVIDfast','cmd':SVIDCMD.FAST,  'data':0x00,'delay':800},'4' :{'descr':'read status1','cmd':SVIDCMD.GETREG,'data':0x10,'delay':20},
               '5' :{'descr':'read PS   ','cmd':SVIDCMD.GETREG,'data':0x32,'delay':20 },'6' :{'descr':'read VID    ','cmd':SVIDCMD.GETREG,'data':0x31,'delay':20},
               '7' :{'descr':'set PS    ','cmd':SVIDCMD.PS,    'data':0x03,'delay':20 },'8' :{'descr':'read PS     ','cmd':SVIDCMD.GETREG,'data':0x32,'delay':20},
               '9' :{'descr':'setVIDfast','cmd':SVIDCMD.FAST,  'data':0x82,'delay':800},'10':{'descr':'read status1','cmd':SVIDCMD.GETREG,'data':0x10,'delay':20},
               '11':{'descr':'read PS   ','cmd':SVIDCMD.GETREG,'data':0x32,'delay':20 },'12':{'descr':'read VID    ','cmd':SVIDCMD.GETREG,'data':0x31,'delay':20}}

#-------------------------------------------------------------------------
# Main program

setup_tool_display(test_rail,start_voltage,test_current)
setup_initial_voltage_and_current(test_rail,start_voltage,test_current)
vResult=setup_and_run_vectors(test_rail,vector_dict)
output_screen(vResult,vector_dict)
shutdown_tool(test_rail)

#
#--------------------------------------------------------------------------
#SVID Vector / Expected Output
    #vector_api.CreateSimpleVector(1, rail.VRAddress,SVIDCMD.GETREG,0x31,20 ) # Read start VID
    #vector_api.CreateSimpleVector(2, rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read start power state
    #vector_api.CreateSimpleVector(3, rail.VRAddress,SVIDCMD.FAST,  0x00,800) # SetVIDFast to 00h  - NOTE Alert# should remain high
    #vector_api.CreateSimpleVector(4, rail.VRAddress,SVIDCMD.GETREG,0x10,20 ) # get Status1 - should be VR_Settled
    #vector_api.CreateSimpleVector(5, rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read end power state register
    #vector_api.CreateSimpleVector(6, rail.VRAddress,SVIDCMD.GETREG,0x31,20 ) # Read end VID
    #vector_api.CreateSimpleVector(7, rail.VRAddress,SVIDCMD.PS,    0x03,20 ) # Set PS to 3 - should ACK
    #vector_api.CreateSimpleVector(8, rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read PS - should read PS0
    #vector_api.CreateSimpleVector(9, rail.VRAddress,SVIDCMD.FAST,  0x82,800) # SetVIDfast to 900mV
    #vector_api.CreateSimpleVector(10,rail.VRAddress,SVIDCMD.GETREG,0x10,20 ) # get Status1 - should be VR_Settled
    #vector_api.CreateSimpleVector(11,rail.VRAddress,SVIDCMD.GETREG,0x32,20 ) # Read PS - should read PS0
    #vector_api.CreateSimpleVector(12,rail.VRAddress,SVIDCMD.GETREG,0x31,20 ) # Read end VID




