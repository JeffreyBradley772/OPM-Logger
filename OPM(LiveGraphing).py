
import sys
import os
import inspect
# Import the .NET Common Language Runtime (CLR) to allow interaction with .NET
import clr
import numpy as np
import time
from datetime import datetime, date
#import serial_rx_tx
from itertools import count
import matplotlib.pyplot as plt
import matplotlib.axes as axis



now = datetime.now()
 
print("now =", now)
# dd/mm/YY H:M:S
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
print("date and time =", dt_string)

print ("Python %s\n\n" % (sys.version,))

strCurrFile = os.path.abspath (inspect.stack()[0][1])
print(inspect.stack()[0][1])

print ("Executing File = %s\n" % strCurrFile)

# Initialize the DLL folder path to where the DLLs are located
strPathDllFolder = os.path.dirname (strCurrFile)
print ("Executing Dir  = %s\n" % strPathDllFolder)

# Add the DLL folder path to the system search path (before adding references)
sys.path.append (strPathDllFolder)

# Add a reference to each .NET assembly required
clr.AddReference (r"C:\Program Files\Newport\Newport USB Driver\Bin\UsbDllWrap")

# Import a class from a namespace
from Newport.USBComm import *
from System.Text import StringBuilder
from System.Collections import Hashtable
from System.Collections import IDictionaryEnumerator

# Call the class constructor to create an object
oUSB = USB (True)

# Discover all connected devices
bStatus = oUSB.OpenDevices (0, True)


    

if (bStatus) :
    oDeviceTable = oUSB.GetDeviceTable ()
    nDeviceCount = oDeviceTable.Count
    print ("Device Count = %d" % nDeviceCount)

    # If no devices were discovered
    if (nDeviceCount == 0) :
        print ("No discovered devices.\n")
        time.sleep(35)
        print('\n Restart Check at: ' + dt_string)
        os.execv(sys.executable,['py'] + ['opm.py'])

        
    else :
        oEnumerator = oDeviceTable.GetEnumerator ()
        strDeviceKeyList = np.array ([])

        # Iterate through the Device Table creating a list of Device Keys
        for nIdx in range (0, nDeviceCount) :
            if (oEnumerator.MoveNext ()) :
                strDeviceKeyList = np.append (strDeviceKeyList, oEnumerator.Key)

        print (strDeviceKeyList)
        print ("\n")

        strBldr = StringBuilder (64)

        # Iterate through the list of Device Keys and query each device with *IDN?
        for oDeviceKey in strDeviceKeyList :
            strDeviceKey = str (oDeviceKey)
            print (strDeviceKey)
            strBldr.Remove (0, strBldr.Length)
            nReturn = oUSB.Query (strDeviceKey, "*IDN?", strBldr)
            print ("Return Status = %d" % nReturn)
            print ("*IDN Response = %s\n" % strBldr.ToString ())

            now = datetime.now()
            dt_string = now.strftime("%m-%d-%Y %H-%M-%S")            

            #Allows us to constantly see if a signal reading is being returned.
            print("Readings Started " + dt_string)
            
        #Begins File Creation And Writing
        base_dir = os.getcwd()

        chartTime = []
        chartOPM = []
        AvgOPM = []
        chartTime2 = []
        plt.plot(chartTime,chartOPM)
        plt.plot(chartTime2,AvgOPM)
        
        while True:

            os.chdir(base_dir)
            
            now = datetime.now()
            
            dt_string = now.strftime("%m-%d-%Y %H-%M-%S")

            today = date.today()
            todays_date = today.strftime( "%m-%d-%Y" )

            if not os.path.exists( strPathDllFolder + "\\Data\\" + todays_date ):
                os.mkdir( strPathDllFolder + "\\Data\\" + todays_date )
                print( "New Folder created : ", todays_date, "\n\n" )

            os.chdir( strPathDllFolder + "\\Data\\" + todays_date )
            
            #fileName = (dt_string)
            file = open('bOPM Readings For--'+dt_string+'.txt','w')

            os.chdir(base_dir)
            
            initialTime = time.time()
            global elapsedTime

            elapsedTime = 0
            
            lastReading = ''
            #lastLastReading = ''

            #Part of FailSave Method, incase readings start to repeat more then 10 times
            consec_count = 0
            
            #Keeps Track of Two Hour Batch File (7200 Seconds)
            
            while elapsedTime <= 7200:
                
                #pauses script for just under a second because we only want readings written every second,
                #after timing the script run time, 0.97 seconds appeared the best length to pause for.
                time.sleep(0.78)
                
                #query's OPM
                nReturn = oUSB.Query(strDeviceKey, "PM:DPower?", strBldr)
                
                
                reading = strBldr.ToString()
                
                
                #writes  reading to file
                t = datetime.now()
                dt_string = t.strftime("%m/%d/%Y %H:%M:%S")
                print(dt_string, "     ", reading)
                file.write(reading + ' , ' + dt_string + '\n')

                
                #checks time for loop
                currentTime = time.time()
                elapsedTime = (currentTime - initialTime)

                
                #Fail-Save for OPM reinitializations
                if reading == lastReading and float(reading) > -70:
                    consec_count +=1
                else:
                    consec_count = 0

                if consec_count == 10:
                    file.close()
                    print('Consecutive Count =', consec_count, '\n')
                    print('\n Restart Check at: ' + dt_string)
                    time.sleep(35)
                    os.execv(sys.executable,['py'] + ['opm.py'])
                    
                #lastLastReading = lastReading
                lastReading = reading
                

                
                

                #CurrentTimeStr = t.strftime('%H:%M:%S')
                
                cTIME  = t.strftime("%H:%M:%S")
                if len(reading) <= 1 or cTIME == '00:00:00':
                
                    file.close()
                    time.sleep(35)
                    print('\n Restart Check at: ' + dt_string)
                    os.execv(sys.executable,['py'] + ['opm.py'])

                t = datetime.now()
                chartTime.append(t)
                chartOPM.append(float(str(reading)))

                if len(chartTime)>20:
                    Avg = sum(chartOPM[-1:-20:-1])/20
                    AvgOPM.append(Avg)
                    chartTime2.append(t)

                
                
                if len(chartTime) > 3600:
                    chartTime.pop(0)
                    chartOPM.pop(0)

                if len(AvgOPM) > 3600:
                    AvgOPM.pop(0)
                    chartTime2.pop(0)

                plt.cla()

                plt.plot(chartTime, chartOPM, label = 'OPM (dbm)',linestyle = 'solid', color ='r')
                plt.plot(chartTime2, AvgOPM, label = 'Avg OPM (dbm)',linestyle = 'solid', color ='b')
                plt.legend(loc = 'upper left')
                plt.grid(True)
                plt.tight_layout()
                plt.ylim(-60,2)                
                plt.xlabel('Datetime', fontsize = 10)
                plt.ylabel('OPM (dbm)', fontsize = 10)
                plt.title('Real Time OPM Readings')
                plt.pause(0.05)
                plt.ion()
                
                plt.show()
                
                
                

                
                
                
                
            plt.show()
            #Closes Previous Two Hour Batch File    
            file.close()
            #Fail-Save
            print('\n Restart Check at: ' + dt_string)
            os.execv(sys.executable,['py'] + ['opm.py'])
                
            
            
            
            

            
else :
    print ("\n***** Error:  Could not open the devices. *****\n\nCheck the log file for details.\n")

# Shut down all communication
oUSB.CloseDevices ()
print ("Devices Closed.\n")
