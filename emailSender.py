import smtplib
import time
import pandas as pd
from getpass import getpass

filepath = raw_input("contacts filepath: ")
data = pd.read_excel (filepath)
df = pd.DataFrame(data, columns= ['Nome', 'Email'])
i = 0

fromAddr = "leduardoferreira187@gmail.com"
toAddr = "luiz.ferreira.carvalho@hotmail.com"
msg = raw_input("message: ")


# smtp login
username = raw_input("email: ")
pswd = getpass("password: ")
subject = raw_input("subject: ")
server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
server.login(username, pswd)

while i < len(df.index):
    time.sleep(1)
    print("Email sent to " + df.iloc[i]['Nome'] + "<" + df.iloc[i]['Email'] + ">")
    toAddr = df.iloc[i]['Nome'] + "<" + df.iloc[i]['Email'] + ">"
    header = 'To:' + toAddr + '\n' + 'From: ' + fromAddr + '\n' + 'Subject:' + subject + ' \n'
    msg = header + msg
    server.sendmail(fromAddr, toAddr, msg)
    i+=1
server.quit()