import re, smtplib
from configparser import ConfigParser
 
def send_email(subject, body):
    cfg = ConfigParser()
    try:
        cfg.read("email.ini")
        server = cfg.get("smtp", "server")
        addr = cfg.get("smtp", "addr")
        apppass = cfg.get("smtp", "apppass")
        sendOK = cfg.get("smtp", "sendOK", fallback = "no")
    except:
        print("Ошибка чтения настроек email")
        return
    if re.search("yes", sendOK, flags = re.IGNORECASE) == None and re.search(r"\bOK\b", subject, flags = re.IGNORECASE) != None:
        return
    message = f"From: {addr}\nTo: {addr}\nSubject: {subject}\n\n{body}"
    server = smtplib.SMTP_SSL(server)
##    server.set_debuglevel(1)
    server.ehlo(addr)
    server.login(addr, apppass)
    server.sendmail(addr, addr, message)
    server.quit()
 
## only ASCII chars allowed
## только английский текст
##send_email("EGRN bot test subject", "EGRN bot test OK")
