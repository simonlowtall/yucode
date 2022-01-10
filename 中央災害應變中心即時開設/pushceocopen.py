# -*- coding: utf-8 -*-
import urllib.request 
import json
import xml.dom.minidom
import pymysql
import datetime
#----------------------setting----------------
def readSettingXmlText(filename, tag):
    dom = xml.dom.minidom.parse(filename)
    root = dom.documentElement
    text = dom.getElementsByTagName(tag)[0].childNodes[0].data
    return text
def strtofulltimeCT(strtime):
    if strtime is None:
      return strtime
    else:
      time = strtime.split(' ')
#      print (time)
      addhour = 0
      if(time[1] == u'下午'):
          addhour = 12
      ymd = time[0].split('/')
      hms = time[2].split(':')
      if(int(hms[0])+addhour == 24):
          result = datetime.datetime(int(ymd[0]),int(ymd[1]),int(ymd[2]),0,int(hms[1]),int(hms[2]))
      else:
          result = datetime.datetime(int(ymd[0]),int(ymd[1]),int(ymd[2]),int(hms[0])+addhour,int(hms[1]),int(hms[2]))
      return result
#--------------------main------------------
myapi = readSettingXmlText('setting.xml','API')
myhost = readSettingXmlText('setting.xml','IP')
myuser = readSettingXmlText('setting.xml','UserName')
mypasswd = readSettingXmlText('setting.xml','PassWord')
date = datetime.datetime.now()
j = urllib.request.urlopen(myapi)
j_obj = json.loads(j.read().decode('utf-8'))
if (j.getcode() == 200):
    insertdatas = []
    for i in j_obj['EEMResp']['OpenCaseInfoList']['openCaseInfo']:
      caseCode = i['caseCode']
      caseName = i['caseName']
      caseStartTime = i['caseStartTime']
      disName = i['disName']
      caseEndTime = None
      eocName = i['eocName']
      openTier = i['openTier']
      openCaseStatusInfo = i['openCaseStatusInfo']
      insertdata = (caseCode,caseName,eocName,int(openTier),caseStartTime,caseEndTime,disName)
      insertdatas.append(insertdata)
      for eoc in openCaseStatusInfo:
#        caseName = eoc['caseName']
        caseStartTime = eoc['prjStime']
        caseEndTime = eoc['prjEtime']
        eocName = eoc['eocName']
        openTier = eoc['opeLv']
        if openTier is None:
          continue
        else:
          insertdata = (caseCode,caseName,eocName,int(openTier),strtofulltimeCT(caseStartTime),strtofulltimeCT(caseEndTime),disName)
          insertdatas.append(insertdata)
    for x in insertdatas:
      print (x)
    try:
      db = pymysql.connect(host=myhost, user=myuser, passwd=mypasswd, db=str('dg_'+str(date.year)), charset = 'utf8')
      cursor = db.cursor()
      insert_sql = ('insert into ceocopen (caseCode,caseName,eocName,openTier,caseStartTime,caseEndTime,disName) values (%s,%s,%s,%s,%s,%s,%s) on  DUPLICATE key update openTier = values(openTier),caseStartTime = values(caseStartTime),caseEndTime = values(caseEndTime)')
      cursor.executemany(insert_sql,insertdatas)
      db.commit()
      print ('-----Insert Finish-----')
      db.close()
    except pymysql.Error as e:
      print ('Open Mysql Fail')
else:
    print("Web service: " + myapi+ " Connection fail\n")