# -*- coding: utf-8 -*-
"""
Created on Sat Aug 22 01:55:36 2015

@author: RAHUL JAIN
"""

import socket
import win32com.client
def Main():
    HOST = ''
    PORT = 5700
    dtl = win32com.client.Dispatch("DTL.Thermodynamics.Calculator")
    dtl.Initialize()
    serversocket = socket.socket(socket.AF_INET,socket.SOCK_STREAM)
    serversocket.bind((HOST,PORT))
    serversocket.listen(2)
    print('Server Listening.....')
    while True:
        connsocket, addr = serversocket.accept()
        print('Connection from',addr)
        if True:
          data = connsocket.recv(4096)
          if not data: break
          strdata = data.decode()
          splitdata = strdata.split(',')
          T = float(splitdata[6])
          P = float(splitdata[7])
          X1 = float(splitdata[8])
          X2 = float(splitdata[9])
          Property = dtl.CalcProp(splitdata[0],splitdata[1],splitdata[2],splitdata[3],[splitdata[4],splitdata[5]],T,P,[X1,X2])
          PropStr = str(Property[0])
          connsocket.send(PropStr)
        else:
          connsocket.close()
        
        
    serversocket.close()
   
if __name__ == '__main__':
    Main()   