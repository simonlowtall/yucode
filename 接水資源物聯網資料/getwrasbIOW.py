import json
import xml.dom.minidom
import requests
import urllib.request
import datetime
import MySQLdb
from oauthlib.oauth2 import BackendApplicationClient
from requests_oauthlib import OAuth2Session

def readSettingXmlText(filename, tag):
	dom = xml.dom.minidom.parse(filename)
	root = dom.documentElement
	text = dom.getElementsByTagName(tag)[0].childNodes[0].data
	return text
#------------------開始----------------------------
tobj = datetime.datetime.now()
myhost = readSettingXmlText('setting.xml','IP')
myuser = readSettingXmlText('setting.xml','UserName')
mypasswd = readSettingXmlText('setting.xml','PassWord')
client_id = readSettingXmlText('setting.xml','client_id')
client_secret = readSettingXmlText('setting.xml','client_secret')
token_url = readSettingXmlText('setting.xml','token_url')
iow_url = readSettingXmlText('setting.xml','iow_url')
insertdatas = []
#------------------token驗證---------------------
client = BackendApplicationClient(client_id=client_id)
oauth = OAuth2Session(client=client)
token = oauth.fetch_token(token_url,client_id=client_id,client_secret=client_secret)
client = OAuth2Session(client_id, token=token)
#------------------抓物理碼---------------------
db = MySQLdb.connect(host=myhost, user=myuser, passwd=mypasswd, db=str('y'+str(tobj.year)+'_display'), charset = 'utf8')
cursor = db.cursor()
cursor.execute('select * from base_iow')
myresult = cursor.fetchall()
for iow in myresult:
  url = iow_url + iow[0]
  r = client.get(url)
  data = json.loads(r.text)
  time = data['TimeStamp']
  value = data['Value']
  if value < 0:
    value  = 0
  code = '0000000000000'+iow[5]+'000'+iow[2]+iow[3]
  insertdata = (time,code,value)
  insertdatas.append(insertdata)
  print (insertdata)
insert_sql = ('insert ignore into data (inittime,datacode,value) values (%s,%s,%s)')
cursor.executemany(insert_sql,insertdatas)
db.commit()
print ('insert database complete')
db.close()