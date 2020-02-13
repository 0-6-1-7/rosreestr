#!/usr/bin/env python3
import cgi
from recognize import recognize

print("Content-type: text/html")
print("Access-Control-Allow-Origin: *")
print()
imgData = cgi.FieldStorage().getfirst("imgData", "")
###captcha = recognize(..., True) if you want to save incoming captcha images
captcha = recognize(imgData.replace(".","+").replace("_","/").replace("-","="), False)
print(captcha)
print()
