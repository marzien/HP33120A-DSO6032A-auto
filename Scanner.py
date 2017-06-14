'''
Changing HP 33120A generator frequency and autosaving UltraLab saftware.
UltraLab files: freq.uld, freq.uld at D:\Mariaus\Python Scripts
This program dir: D:\Mariaus\Python Scripts\HP su UltraLab
'''

import RS232_instrument
import DSO6032A
import autoUL
import numpy
import time
import win32api
from pywinauto.application import Application
import visa
import pyautogui

path_to_wavegen = "ASRL4::INSTR"
path_to_osc = "USB0::0x0957::0x1732::MY55270124::INSTR"
wavegen = RS232_instrument.agilent33120A(path_to_wavegen)

fStart = 10000 # Hz
fEnd = 70000   # Hz
fStep = 500    # Hz
freq = numpy.arange(fStart, fEnd+fStep, fStep)

for num in freq:

    #RUN osciloscope for measurement
    rm = visa.ResourceManager()
    my_Osc = rm.open_resource(path_to_osc)
    my_Osc.write(":RUN")
    time.sleep(1)
    
    #Function generator parameters
    print("Frequency (kHz):",num/1000)
    wavegen.control("LOCAL")
    wavegen.loadImpedance(50) #50 standard
    wavegen.mode("sin")
    wavegen.voltage(10.0)# x2 real
    wavegen.frequency(num)
    wavegen.burstCount(50)
    wavegen.burstOn(True)
    wavegen.extTig(True)
    time.sleep(1) #M-sequence 5

    #UltraLab automatization
    autoUL.measure(num)
    time.sleep(1)

    #save osciloscope screen to *.csv using DSO6023A.exe
    # at C:\scope\data
    pyautogui.keyDown('alt') #New file
    pyautogui.press('F')
    pyautogui.press('N')
    pyautogui.keyUp('alt')
    
    pyautogui.click(44, 1130) #Execute DSO6023A.exe
    my_Osc.write(":RUN")
    app = Application().start("D:\\DSO6023A.exe", wait_for_idle = False)
    app.Window_().TypeKeys(str(num)+'{ENTER}', pause = 0.1)
    time.sleep(3)
    pyautogui.click(260, 1129) #Stop
    
print("Finished!!!")
    
