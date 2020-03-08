#!/usr/bin/python

import urllib3
import os.path
import syslog

http = urllib3.PoolManager();
oldIP = '';
ipfilename = '/etc/dnsomatic/currentip';
authentication = 'Basic Y8AJja8d89ad82347SJFAHF873adhs7aAJ878'
hostname = 'www.example.com'

syslog.syslog(syslog.LOG_INFO, "dns-o-matic processing started")

hd = {'Authorization':authentication,
'Content-Type':'application/x-www-form-urlencode'}

if(os.path.isfile(ipfilename)):
    ipfile = open(ipfilename,"r")
    oldIP = ipfile.readline().rstrip();
    ipfile.close() 
else:
    ipfile = open(ipfilename,"a")
    ipfile.close()
    
r = http.request("GET", "http://myip.dnsomatic.com/")
currentIP = r.data
if currentIP != oldIP:
   ipfile = open(ipfilename,"w")
   ipfile.write(currentIP)
   ipfile.close()

   flds = {'hostname':hostname,
   'myip':currentIP,
   'wildcard':'NOCHG',
   'mx':'NOCHG',
   'backmx':'NOCHG'}

   r = http.request("POST", "http://updates.dnsomatic.com/nic/update/", fields=flds, headers=hd)

   if(r.data.find("good") != -1):
      syslog.syslog(syslog.LOG_INFO, "dns-o-matic update successful! [" + r.data + "]" ) 
   else:
      syslog.syslog(syslog.LOG_INFO, "dns-o-matic update failed! [" + r.data + "]")  
else:
   syslog.syslog(syslog.LOG_INFO, "IP address has not changed, exiting")
