# -*- coding: utf-8 -*-
"""
Created on Wed Aug 19 03:13:38 2015

@author: RAHUL JAIN
"""


import socket
import win32com.client
def Main():
    HOST = ''
    PORT = 5000
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
          P = float(splitdata[1])
          T = float(splitdata[2])
          X1 = float(splitdata[5])
          X2 = float(splitdata[6])
          PTFlash = dtl.PTFlash(splitdata[0],0,P,T,[splitdata[3],splitdata[4]],[X1,X2])
          ptfl = str(PTFlash[1][0]) + " " + str(PTFlash[2][0]) + " " + str(PTFlash[2][1])
          connsocket.send(ptfl)
        else:
          connsocket.close()
        
        
    serversocket.close()
   
if __name__ == '__main__':
    Main()   