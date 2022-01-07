#-*- coding: utf8 -*- 
import sys , os
import string
import datetime
import time
import xml.dom.minidom
import sys
import urllib.request
import MySQLdb
import json
def readSettingXmlText(filename, tag):
    dom = xml.dom.minidom.parse(filename)
    root = dom.documentElement
    text = dom.getElementsByTagName(tag)[0].childNodes[0].data
    return text
def strtofulltimeCT(strtime):
    time = strtime.split('T')
    #print time
    result = time[0][0:4]+'/'+time[0][5:7]+'/'+time[0][8:10]+' '+time[1][0:5]
    return result
myhost = readSettingXmlText('setting.xml','IP')
myuser = readSettingXmlText('setting.xml','UserName')
mypasswd = readSettingXmlText('setting.xml','PassWord')
insertdatas = []
#--------------------main------------------
x = datetime.datetime.now()
j = urllib.request.urlopen('https://opendata.cwb.gov.tw/fileapi/v1/opendataapi/O-A0002-001?Authorization=CWB-DDD472D9-32BB-4DFC-995A-91BE320C3E6B&format=JSON')
j_obj = json.loads(j.read())
#print j_obj
if (j.getcode() == 200):
  for data in j_obj['cwbopendata']['location']:
    if data['time']['obsTime'].split('T')[1][3:5] == '00':
      DateTime = strtofulltimeCT(data['time']['obsTime'])
      Rain_Station_ST_NO = data['stationId']
      Rainfall = data['weatherElement'][1]['elementValue']['value']
      if Rainfall in ('-998.00','-999.00'):
        Rainfall = '0.00'
      print (Rain_Station_ST_NO,DateTime,Rainfall)
      insertdata = (Rain_Station_ST_NO,DateTime,Rainfall)
      insertdatas.append(insertdata)
    else:
      print (x,u'尚無整點資料')
      sys.exit()
  try:
      db = MySQLdb.connect(host=myhost, user=myuser, passwd=mypasswd, db=str('dg_'+str(x.year)), charset = 'utf8')
      cursor = db.cursor()
      insert_sql = ('insert ignore into rt_rainfall (Rain_Station_ST_NO,DateTime,Rainfall) values (%s,%s,%s)')
      cursor.executemany(insert_sql,insertdatas)
      db.commit()
      print ('-----Finish-----')
      db.close()
  except MySQLdb.Error as e:
      print ('Open Mysql Fail')
else:
    print("Web service: " + myhttpconnection+ " Connection fail\n")