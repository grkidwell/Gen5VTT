# ------------------------------------------------------------------------------------
# ----Author - G. Kidwell
# 
# 
# ---Modified from Intel DC_Load_Example
# 
#--------------------------------------------------------------------------------------


import clr, os, sys, time

sys.path.append(os.path.join(sys.path[0],'intel'))

from Common.APIs import DataAPI, DisplayAPI, MeasurementAPI, MeasurementItemsAPI, GeneratorAPI
from Common.Models import TriggerModel,FanModel,RawDataModel,DataModel,DisplayModel
from Common.Enumerations import HorizontalScale, ScoplessChannel, Cursors,PowerState,Transition,Protocol,RailTestMode
#from System import String, Char, Int32, UInt16, Boolean, Array, Byte, Double

generator_api = GeneratorAPI();  rails = generator_api.GetRails();   data_api = DataAPI()
display_api = DisplayAPI();      measurement_api = MeasurementAPI();  measurement_items_api = MeasurementItemsAPI()

def ddefine_rail_data(testrail):
    try:
        rail = [r for r in generator_api.GetRails() if r.Name==testrail][0]
    except:
        print('Test failed because the test rail could not be found. Please check the rail name.')
        exit()
    return rail

def define_rail_data(testrail):
    rail = None

    for r in list(generator_api.GetRails()):
        if r.Name == testrail:
            rail = r
    if rail is None:
        print('Test failed because the test rail could not be found. Please check the rail name.')
    
    return rail



def setup_tool_display(testrail):
    data_api.SetFanSpeed(10)
    display_api.SetHorizontalScale(HorizontalScale.Scale10us)
    data_api.SetTrigger(4, 0, 0, 0, 0, False)
    display_api.Ch1_2Rail(testrail, True)  # True enables autoscaling in display, False disables
    display_api.SetChannel(ScoplessChannel.Ch1, True)
    display_api.SetChannel(ScoplessChannel.Ch2, True)

def setup_initial_voltage(testrail,testvoltage):
    generator_api.Generator1SVSC(testrail, 0, False)
    generator_api.SetVoltageForRail(testrail, testvoltage, Transition.Fast)

def enable_measurements(testrail):
    measurement_api.MeasureCurrentMean(testrail)
    measurement_api.MeasureVoltageMean(testrail)
    measurement_api.MeasureVoltageAmplitude(testrail)
    measurement_api.MeasurePersistentVoltageMinMax(testrail)
    measurement_api.MeasurePersistentCurrentMinMax(testrail)

def scale_imon(imon_hex_value,iccmax):
    return int(imon_hex_value,16)/0xff*iccmax

def read_imon(raildata):
    svid_transaction = None
    if raildata is not None:
        svid_command = 0x7  # Get Reg
        imon_register_address = 0x15
        data_api.SetSvidCmdWrite(raildata.VRAddress, svid_command, imon_register_address, raildata.SVIDBus)
        time.sleep(0.25)
        svid_transaction = data_api.GetSvidData()
    svid_data=0
    if svid_transaction is not None:
        svid_data=svid_transaction.SVRData
    return svid_data
    
def sweep_load_current(raildata,testrail,startcurrent,endcurrent,increment,iccmax):
    current_offset=0.25
    #test_currents = [1,5,10,15,50,100]  # A
    for current in range(startcurrent, endcurrent+increment, increment):
        test_current=current+current_offset
        #test_current=current
        generator_api.Generator1SVSC(testrail, test_current, True)
        time.sleep(1)
        measurement_api.ResetPersistentCurrentMeasurement(testrail)
        measurement_api.ResetPersistentVoltageMeasurement(testrail)
        time.sleep(1)
        voltage_measure = measurement_items_api.GetVoltageMeanOnce(testrail)[0].Value
        current_measure = measurement_items_api.GetCurrentMeanOnce(testrail)[0].Value
        imon = read_imon(raildata)

        #print("{:.3f} \t {:.3f} \t {}".format(current_measure, voltage_measure, imon))
        print(f"{current_measure}\t{voltage_measure}\t{imon}")


# End Of imports
print("##############################################################################")
print("------------------------------------------------------------------------------")
print("                              DC LOAD LINE                                    ")
print("------------------------------------------------------------------------------")
print("##############################################################################")
print("")

test_rail = "VCCCORE"
#test_currents = [0, 2]  # A
test_voltage = 0.9  # V
icc_max = 160
start_current = 0
end_current = 50
increment = 5

rail_data = define_rail_data(test_rail)
print(rail_data.Name)

setup_tool_display(test_rail)
setup_initial_voltage(test_rail,test_voltage)
enable_measurements(test_rail)

print("Begining Test")



print("")
print('  .............................................................................................................')
print('  Rail Current \t Rail Voltage\t IMON')
print('  .............................................................................................................')

sweep_load_current(rail_data,test_rail,start_current,end_current,increment,icc_max)

generator_api.Generator1SVSC(test_rail, 0, False)
print("---------------------------------------------------------------------------------------------------------------")
print("")
print("Test complete")
