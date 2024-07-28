import pandas as pd
import datetime
from smtplib import SMTP_SSL as SMTP
from email.mime.text import MIMEText as MT
from email.mime.multipart import MIMEMultipart as MM
import sys

def sendEmail(to, sub, msg ):
    SMTPserver = ''
    sender =     'teamtrio2003@gmail.com'
    destination = [to]

    USERNAME = ""
    PASSWORD = ""
    
    try:
        msg = MM()
        msg['Subject']=sub
        msg['From']   = sender # some SMTP servers will do this automatically, not all
        
        HTML = """
        <html>
        <body>
        <h1> Happy Birthday! </h1>
        <img src="bday_card.jpg" alt = "image" width="400" height="360">
        <p>Happy birthday to you 🎂🎈Another year older, another year wiser, and another year ready to take on the world 🍰😊
        </p>
        </body>
        </html>
        """
        
        MTobj = MT(HTML,"html")
        msg.attach(MTobj)

        conn = SMTP(SMTPserver)
        conn.set_debuglevel(False)
        conn.login(USERNAME, PASSWORD)
        try:
            conn.sendmail(sender, destination, msg.as_string())
        finally:
            conn.quit()
    except:
        sys.exit( "mail failed; %s" % "CUSTOM_ERROR" ) # give an error message


if __name__ == "__main__":
    
    df = pd.read_excel("data_new.xlsx")
    # print(df)
    today = datetime.datetime.now().strftime("%d-%m")
    yearNow = datetime.datetime.now().strftime("%Y")
    # print(type(today))
    writeInd = []
    # print(df)
    for index, item in df.iterrows():
        # print(">>>>",index, item['Birthday'])
        bday = item['Birthday'].strftime("%d-%m")
        # print(bday) 
        if(today == bday) and yearNow not in str(item['Year']):
            
            sendEmail(item['Email'], "Happy Birthday", item['Dialogue'] ) 
            writeInd.append(index)

    # print(writeInd)
    for i in writeInd:
        yr = df.loc[i, 'Year']
        df.loc[i, 'Year'] = str(yr) + ', ' + str(yearNow)
        # print(df.loc[i, 'Year'])

    # print(df) 
    df.to_excel('data_new.xlsx', index=False)   