# Create your views here. 
# -*- coding: utf-8 -*-
import MySQLdb
import xlwt
from datetime import datetime

wbk=xlwt.Workbook(encoding='utf-8')
sheet=wbk.add_sheet('sheet 1')
sheet.write(0,0,'Uid')
sheet.write(0,1,u'姓名')
sheet.write(0,2,'password')
sheet.write(0,3,'Email')
sheet.write(0,4,'Date')
row=1
fmt='YYYY-MM-DD hh:mm:ss'
try:
    conn=MySQLdb.connect(host="localhost",user='root',passwd="",db="personalblog_db",charset='utf8')
    cur=conn.cursor()
    count=cur.execute('select * from users')
    print 'there has %s rows record' % count
    results=cur.fetchmany(count)
    for uid,uname,upassword,uemail,udate in results:
        sheet.write(row,0,uid)
        sheet.write(row,1,uname)
        sheet.write(row,2,upassword)
        sheet.write(row,3,uemail)
        style = xlwt.XFStyle()
        style.num_format_str = fmt
        sheet.write(row,4,udate,style)
        row+=1
    wbk.save('users.xls')
    cur.close()
    conn.commit()
    conn.close()
except MySQLdb.Error,e:
    print "Mysql Error %d: %s"%(e.args[0],e.args[1])
