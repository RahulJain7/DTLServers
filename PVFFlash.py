# -*- coding: utf-8 -*-
"""
Created on Mon Sep 07 17:05:31 2015

@author: RAHUL JAIN
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Aug 19 03:13:38 2015

@author: RAHUL JAIN
"""


import socket
import win32com.client
def Main():
    HOST = ''
    PORT = 7000
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
          Nc = int(splitdata[3])
          No = 4+Nc
          P = float(splitdata[1])
          VF = float(splitdata[2])
          Comp = splitdata[4:No]
          Xstr = splitdata[No:len(splitdata)]
          X = [float(i) for i in Xstr]
          PVFlash = dtl.PVFFlash(splitdata[0],0,P,VF,Comp,X)
          ptfl = " " + str(PVFlash[2][0]) + " "
          if Nc>2:
           for j in range(3,Nc+1):
             ptfl = ptfl + str(PVFlash[j][0]) + " "
          ptfl = ptfl + PVFlash[Nc+2][0]
          connsocket.send(ptfl)
        else:
          connsocket.close()
        
        
    serversocket.close()
   
if __name__ == '__main__':
    Main()   