# -*- coding: utf-8 -*-
"""
Created on Tue Aug 18 18:33:58 2015

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
          data = connsocket.recv(1024)
          if not data: break
          strdata = data.decode()
          splitdata = strdata.split(',')
          T = float(splitdata[1])
          VapPres = dtl.GetCompoundTDepProp(splitdata[0],"vaporPressure",T)
          VapPresstr = str(float(VapPres))
          connsocket.send(VapPresstr)
        else:
          connsocket.close()
        
        
    serversocket.close()
   
if __name__ == '__main__':
    Main()   