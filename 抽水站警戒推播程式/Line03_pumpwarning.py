# -*- coding: utf-8 -*-
import os
import datetime
import sys
import xml.dom.minidom
import json
import pymysql
import urllib3
import urllib.request
import requests
import ssl

ssl._create_default_https_context = ssl._create_unverified_context


urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
      
def readSettingXmlText(filename, tag):
    dom = xml.dom.minidom.parse(filename)
    root = dom.documentElement
    text = dom.getElementsByTagName(tag)[0].childNodes[0].data
    return text

#read setting
myhttpconnection = readSettingXmlText('setting.xml','url')
abntau = readSettingXmlText('setting.xml','tau')
#read warning
L1 = readSettingXmlText('warninginfo.xml','L1')
L2 = readSettingXmlText('warninginfo.xml','L2')
L3 = readSettingXmlText('warninginfo.xml','L3')
L1text = readSettingXmlText('warninginfo.xml','L1text')
L2text = readSettingXmlText('warninginfo.xml','L2text')
L3text = readSettingXmlText('warninginfo.xml','L3text')


def lineNotifyMessage(token, msg):
    headers = {
        "Authorization": "Bearer " + token, 
        "Content-Type": "application/x-www-form-urlencoded"
    }
    payload = {'message': msg}
    r = requests.post(
        "https://notify-api.line.me/api/notify",
        headers=headers, 
        params=payload)
    return r.status_code
token = '放入你的line notify token'
tobj = datetime.datetime.now()
timestr = datetime.datetime.strftime(tobj, "%Y-%m-%dT%H:%M:%S")
lostdatapushtime = datetime.datetime.strftime(tobj, "%H%M")
print("Current Time =" + datetime.datetime.strftime(tobj, "%Y/%m/%d+%H:%M"))
#-----------------------------------------檢查log檔---------------------------------------------
with open('logfile.log','r') as f:
    readlogdata = f.read()
logtime = readlogdata.split(',')[0]
logwlevel = readlogdata.split(',')[1]
logfile = open('logfile.log','w+')
print (readlogdata)
#-----------------------------------------以下是程式本體--------------------------------------------
request = urllib.request.Request(myhttpconnection, headers={"Accept" : "application/json"})
j = urllib.request.urlopen(request)
if (j.getcode() == 200):
  j_obj = json.loads(j.read().decode('utf-8'))
  insertdatas = []
  realdata = j_obj['Data']
  for key in realdata:
    if key['StationNo'] == '1430H028':
      waterlevel   = key['WaterLevel']
      time   = key['Time']
      print (waterlevel,time)
  if float(L2)>float(waterlevel)>float(L1):
      if int(logwlevel) > 1:
          logfile.write(time+',1')
      elif (tobj - datetime.datetime.strptime(logtime,'%Y-%m-%dT%H:%M:%S')).seconds < int(abntau) and int(logwlevel)>0:
          logfile.write(readlogdata)
      else:
          logfile.write(time+',1')
          lineNotifyMessage(token,L1text)
          print(L1text)
  elif float(L3)>float(waterlevel)>float(L2):
      if int(logwlevel) > 2:
          logfile.write(time+',2')
      elif (tobj - datetime.datetime.strptime(logtime,'%Y-%m-%dT%H:%M:%S')).seconds < int(abntau) and  int(logwlevel) == 2:
          logfile.write(readlogdata)
      else:
          logfile.write(time+',2')
          print(L2text)
          lineNotifyMessage(token,L2text)
  elif float(waterlevel)>float(L3):
      if  (tobj - datetime.datetime.strptime(logtime,'%Y-%m-%dT%H:%M:%S')).seconds < int(abntau) and int(logwlevel) == 3:
          logfile.write(readlogdata)
      else:
          logfile.write(time+',3')
          print(L3text)
          lineNotifyMessage(token,L3text)
  else:
    logfile.write(time+',0')
logfile.close()