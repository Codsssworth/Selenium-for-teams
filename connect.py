import datetime
import time
import pyautogui
import os
import socket
import pygetwindow as gw
import subprocess

pyautogui.FAILSAFE = False


tz=time.tzname

def utime():
    t = time.localtime()
    return t



def isConnected():
  try:
    # reachable
    sock = socket.create_connection(("www.google.com", 80))
    if sock is not None:

      print(sock.type )
      sock.close()

    return True
  except OSError:
    pass
  return False

def connect(hour,min,sec):
    window = gw.getWindowsWithTitle('Microsoft Teams - Google Chrome')
    d= len(window)
    if d==0  and isConnected()==True :

            print("Running Batch at ",hour,":",min,":",sec)
            subprocess.call([r'C:\Users\Bismay\Desktop\team.bat'])

    elif d>0:
        print("Already Running on", window[0])
    else:
        print("OOOOOOOOOOOOoooooo LOST INTERNET")



while True:
    hour = utime().tm_hour
    min = utime().tm_min
    sec=utime().tm_sec

    if hour>=7 and hour <18:

            print("DIAG'NOSING AT :",hour ,":" ,min )
            connect(hour,min,sec)
            print("Quit - out of auto joiner")
            time.sleep(300)
    else:
        print("[",hour,':',min,':',sec,"]Maybe when the time is right :) ")
        time.sleep(60);








