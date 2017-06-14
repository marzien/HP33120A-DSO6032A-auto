import pyautogui
import time
from datetime import datetime

def measure(freq):
    fName = str(freq) 
    #pyautogui.hotkey('ctrl', 'N')
    pyautogui.click(370, 1130)
    pyautogui.keyDown('alt')
    pyautogui.press('F')
    pyautogui.press('N')
    pyautogui.keyUp('alt')
    
    pyautogui.click(44, 1130)# Start
    #startTime = time.ctime()
    #print("Start Time : %s" % time.ctime())
    startTime = datetime.now()
    time.sleep(18)# sleep for measurement
    print("Time of measurement: %s" % (datetime.now() - startTime))
    
    
    pyautogui.click(260, 1129)# Stop
    time.sleep(2)
    
    # Save dealog
    print("Saving as : %s" % fName)
    pyautogui.keyDown('alt')
    pyautogui.press('F')
    pyautogui.press('S')
    pyautogui.keyUp('alt')
    pyautogui.click(561, 746) # file name
    pyautogui.typewrite(fName)
    pyautogui.press('enter')  # press the Enter key
    
    print('Done UltraLab saving.')

if __name__=='__main__':
    measure(freq)  
