import socket
import pyautogui
import pygetwindow as gw
import os
from datetime import datetime
import time
from auto_joiner import browser
import auto_joiner as aj



pyautogui.FAILSAFE = False


# def cmd():
#       os.startfile('C:\Windows\System32\cmd.exe')
#       pyautogui.moveTo(279, 264)
#       pyautogui.typewrite(' cd Desktop\FailS', interval=0.05)
#       pyautogui.press('return')
#       pyautogui.PAUSE = 1.5
#       pyautogui.typewrite('python auto_joiner.py', interval=0.05)
#       pyautogui.press('return')


def isConnected():
  try:
    # reachable
    sock = socket.create_connection(("www.google.com", 80))
    if sock is not None:

      print(sock)
      sock.close()

    return True
  except OSError:
    pass
  return False

def getmeet():
    global mt
    if len(aj.meetings) >= 0:
        print("looking for meetings: ")

        for meeting in aj.meetings:
            print(meeting, len(aj.meetings), "time :", meeting.time_started)
            mt=meeting.time_started
            print("time :",type(mt),mt)

    else:
        print("no current meeting")



if __name__ == '__main__':
    aj.load_config()
    meet=aj.meetings


    while 1:
        global mt
        recconnect=0
        getmeet()
        timestamp = datetime.now()
        print(f"\nStarting Fails[{timestamp:%H:%M:%S}]")


        if isConnected() == True and recconnect == 0:
            print(f"\nStarting Script at : [{timestamp:%H:%M:%S}] ")
            start = aj.main()

            if recconnect == 1:
                print(f"\n[{timestamp:%H:%M:%S}] Restarting due to internet loss")
                time.sleep(5)
                aj.load_config()
                aj.main()
                recconnect = 0

        elif mt!=0:
            print("www")


        else:
            recconnect = 1
            print("Internet lost")















