import win32com.client as win32
import pandas as pd
outlook = win32.Dispatch('Outlook.Application')

filepath = "C:\\Users\\braaten\\Desktop\\Email Automation\\"
excelfile = "C:\\Users\\braaten\\Desktop\\Email Automation\\Email Data.xlsx"
df = pd.read_excel(excelfile).fillna('')

firstname = df["First Name"].tolist()
send_to_email = df["email"].tolist()
org = df["Organization"].tolist()
                 
#Sender email addresses - don't end with semicolon
me = "homer@snpp.org"
myboss = "smithers@snpp.org"
ourCEO = "mrburns@snpp.org"

##Attachments
filename1 = "MBAOAIFE Letter of Request "
suffix = ".pdf"

#Message subject and body
subject = "We ask for your support of the Montgomery Burns Award for Outstanding Acheivement in the Field of Excellence"
body1 = """,
<br><br>
We humbly ask that you renew your donation in the amount of $1,000,000,000,000 in the usual denomination of $10,000 bills with all the presidents having a party.
<br><br>
Best regards,
<br><br>
Ol Gil
<br><br>
PS - Please, I need this!
"""

for i in range (0,len(firstname)):
    new_mail = outlook.CreateItem(0)
    
    new_mail.SentOnBehalfOfName = myboss
    new_mail.Sender = myboss
    
    #CC and BCC separate with semicolons
    new_mail.To = send_to_email[i]
    new_mail.CC = ourCEO
    new_mail.BCC = me
    
    new_mail.Importance = "2" # 1 is low; 2 is high
    new_mail.DeferredDeliveryTime = 20398.83 #use excel to find numerical value for time and date; this float corresponds to 8pm on a redletterdate in the history of science, November 5th, 1955
    
    new_mail.Subject = subject 
    new_mail.HTMLBody = "<p>Dear "+firstname[i]+body1
    
    new_mail.Attachments.Add(filepath+filename1+org[i]+suffix)

    new_mail.Save() # Saves file to the draft folder in outlook
    #new_mail.Send() # Sends file to the Outlook outbox


# In[ ]:




