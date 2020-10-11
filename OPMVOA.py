import sys
import os
import inspect
# Import the .NET Common Language Runtime (CLR) to allow interaction with .NET
import clr
import numpy as np
import time
from datetime import datetime, date
import serial
import serial_rx_tx



now = datetime.now()
 
print("now =", now)
# dd/mm/YY H:M:S
dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
print("date and time =", dt_string)

print ("Python %s\n\n" % (sys.version,))

strCurrFile = os.path.abspath (inspect.stack()[0][1])

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

    # If no devices were discovered, re-executes code till it is found.
    if (nDeviceCount == 0) :
        print ("No discovered devices.\n")
        time.sleep(30)
        print('Restart Check at: ' + dt_string)
        os.execv(sys.executable,['python'] + ['opmVOA.py'])

        
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

            
            print("Readings Started " + dt_string)
        base_dir = os.getcwd()
        while True:

            os.chdir(base_dir)
            today = date.today()
            todays_date = today.strftime('%m-%d-%Y')

            if not os.path.exists( strPathDllFolder + "\\Data\\" + todays_date ):
                os.mkdir( strPathDllFolder + "\\Data\\" + todays_date )
                print( "New Folder created : ", todays_date, "\n\n" )

            os.chdir( strPathDllFolder + "\\Data\\" + todays_date )
            
            #VOA Serial Connection
            VOA = serial_rx_tx.SerialPort()
            
            #Setip Com Port For VOA Serial Connection
            VOA.Open('COM7',9600)

            #Desired OPM Level the VOA will adjust to keep the signal at
            DesiredOPMLevel = -20


            #File Name Setup
            now = datetime.now()
            dt_string = now.strftime("%m-%d-%Y %H-%M-%S")
            fileName = (dt_string)
            file = open('OPM Readings For--'+dt_string+'.csv','w')

            os.chdir(base_dir)
            
            
            #Intial Time For Keeping Track of Two hour Files 
            initialTime = time.time()
            elapsedTime = 0

            #Keeps Track of Two Hour Batch File (7200 Seconds)
            lastReading = ''
            consec_count = 0
            
            while elapsedTime <= 7200:
                TTime = time.time()
                TElapsed = 0

                #This loop was implemented so the Algo would query the OPM as fast
                #as it could for adjusting the VOA, but we only want a reading written
                #to the file every second. Once TElapsed reaches one second,
                #the while loop breaks and a reading is written and then it starts over.
                
                while TElapsed < 1:
                    
                    
                    nReturn = oUSB.Query(strDeviceKey, "PM:DPower?", strBldr)
                    
                    
                    reading = strBldr.ToString()
                    #Try and Except Because sometimes the OPM returns a header
                    #instead of a power value and we want to convert that power value into
                    # a float
                    
                    try:
                        #needs to be converted to a string type then a float because it originates
                        #as a 'stringBuilder' type
                        OPM = float(str(reading))
                        
                        
                        if abs(OPM - DesiredOPMLevel) >= 1:
                            if (OPM - DesiredOPMLevel) > 0: #OPM 'higher' or stronger signal than desired
                                #Sets attneuation to difference between OPM level and Desired Level
                                VOA.Send('A+'+ str(int(abs(OPM - DesiredOPMLevel))))
                                #print('reseting')
                            else: #OPM lower or weaker signal then desired, so we need to turn attenuation DOWN
                                #Decreases Attenuation by the difference 
                                VOA.Send('A-'+ str(int(abs(OPM - DesiredOPMLevel))))
                                #print('lowering')
                    except:
                        print("Error, not a number: ", reading)
                        continue
                    
                        
                        
                    #Not Necessary, but helps with importing data into MatLab later-on
                    try:
                        reading = str(float(str(reading)))
                        #print(reading)
                    except:
                        continue

                    TElapsed = (time.time() - TTime)


                    
                     
                #After 1 second is reached, the following writes the most
                #recent reading to a file.
                t = datetime.now()
                dt_string = t.strftime("%m/%d/%Y %H:%M:%S")

                print(dt_string, "   |   ", reading)
                
                file.write(reading + ' , ' + dt_string + '\n')

                

                currentTime = time.time()
                elapsedTime = (currentTime - initialTime)
                

                #CurrentTimeStr = t.strftime('%H:%M:%S') 

                #Fail-Save implemented to make sure a reading is getting received,
                #We noticed the OPM likes to reinitialize from time to time and gives no readings.
                #This Stops the script and reexecutes it to ensure it reconnects to the OPM
                #Once It finishes reinitializing.
            
                if len(reading) <= 1:
                
                    file.close()
                    time.sleep(30)
                    print('Restart Check at: ' + dt_string)
                    os.execv(sys.executable,['py'] + ['opmVOA.py'])

                if reading == lastReading and float(reading) > -70:
                    consec_count +=1
                else:
                    consec_count = 0

                if consec_count == 10:
                    file.close()
                    print('Consecutive Count =', consec_count, '\n')
                    print('\n Restart Check at: ' + dt_string)
                    time.sleep(35)
                    os.execv(sys.executable,['py'] + ['opmVOA.py'])

                lastReading = reading




                #CurrentTimeStr = t.strftime('%H:%M:%S')

                cTIME  = t.strftime("%H:%M:%S")
                if len(reading) <= 1 or cTIME == '00:00:00':

                    file.close()
                    time.sleep(35)
                    print('\n Restart Check at: ' + dt_string)
                    os.execv(sys.executable,['py'] + ['opmVOA.py'])
                
                
                

            #close Two hour batch file before starting a new one    
            file.close()
            
            ##print('\n Restart Check at: ' + dt_string) CHANGE TO ACTUAL FILE NAME
            ##os.execv(sys.executable,['py'] + ['opm.py'])
                
            
            
            
            

            
else :
    print ("\n***** Error:  Could not open the devices. *****\n\nCheck the log file for details.\n")

# Shut down all communication
oUSB.CloseDevices ()
print ("Devices Closed.\n")
