import datetime
import xml.dom.minidom
import MySQLdb
import requests
import json
import ast

def readSettingXmlText(filename, tag):
	dom = xml.dom.minidom.parse(filename)
	root = dom.documentElement
	text = dom.getElementsByTagName(tag)[0].childNodes[0].data
	return text

tobj = datetime.datetime.now()
myhost = readSettingXmlText('setting.xml','IP')
myuser = readSettingXmlText('setting.xml','UserName')
mypasswd = readSettingXmlText('setting.xml','PassWord')
url = readSettingXmlText('setting.xml','url')
login = url+'/user/auth/login'
print (myhost,myuser,mypasswd)
print (login)

headers = {"accept": "application/json","Content-Type": "application/json"}
data = "{\"password\": \"password\",\"tenant\": \"tenant\",\"username\": \"username\"}"
#data內容因安全性因素拿掉了密碼及其他設定
request = requests.post(login,headers=headers,data=data)
if (request.status_code == 201):
  insertdatas = []
  response = request.content.decode()
  token = ast.literal_eval(response)
  #print (token['token'])
  get_rain = url+'/r/W4V1qe/water/rain/getRainList?org_id=109&supplier=2%2C3'
  print (get_rain)
  Authorization = token['token']
  headers = {"accept": "application/json","Authorization": Authorization}
  request = requests.get(get_rain,headers=headers)
  response = request.content.decode()
  #print (response)
  j_obj = json.loads(response)
  for key in j_obj:
    Rain_Station_ST_NO = key['st_no']
    DateTime = key['datatime']
    Rainfall = key['min_10']
    insertdata = (Rain_Station_ST_NO,DateTime,Rainfall)
    insertdatas.append(insertdata)
  print (insertdatas)
  db = MySQLdb.connect(host=myhost, user=myuser, passwd=mypasswd, db='dg_'+str(tobj.year), charset = 'utf8')
  cursor = db.cursor()
  insert_sql = ('insert ignore into rt_rainfall (Rain_Station_ST_NO,DateTime,Rainfall) values (%s,%s,%s)')
  cursor.executemany(insert_sql,insertdatas)
  db.commit()
  print ('insert rt_rainfall complete')
  db.close()
else:
    print("Web service: " + url+ " Connection fail\n")
