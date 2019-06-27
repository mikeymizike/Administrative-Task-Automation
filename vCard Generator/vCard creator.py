import os
import pandas as pd
filepath = "C:\\Users\\braaten\\Desktop\\vCards\\"
excelfile = "C:\\Users\\braaten\\OneDrive - NCQA\\scratch worksheet\\scratch worksheet.xlsx"
df = pd.read_excel(excelfile, sheet_name = "vCards", usecols="A:J", dtype={"zipcode":'str'}).dropna()
last = df["Last Name"].tolist()
first = df["First Name"].tolist()
org = df["Organization"].tolist()
title = df["Title"].tolist()
phone = df["Phone"].tolist()
email = df["email"].tolist()
address = df["address"].tolist()
city = df["City"].tolist()
state = df["State"].tolist()
df['zipcode'] = df['zipcode'].astype(str).str.zfill(5)
zipcode = df["zipcode"].tolist()
country = "United States of America"

for i in range(0,len(last)):
    file = open(filepath+first[i]+" "+last[i]+'.vcf', 'w')
    file.write("BEGIN:VCARD\nVERSION:2.1\nN;LANGUAGE=en-us:\n")
    file.write(last[i]+";"+first[i]+"\n")
    file.write("FN:"+first[i]+" "+last[i]+"\n")           
    file.write("ORG:"+org[i]+"\n")
    file.write("TITLE:"+title[i]+"\n")
    file.write("TEL;WORK;VOICE:"+"("+str(phone[i])[:3]+") "+str(phone[i])[3:6]+"-"+str(phone[i])[6:10]+"\n")
    file.write("ADR;WORK;PREF:;;"+address[i]+";"+city[i]+";"+state[i]+";"+str(zipcode[i])+";"+country+"\n")
    file.write("EMAIL;PREF;INTERNET:"+email[i]+"\n")
    file.write("END:VCARD")
    file.close()
file.close()

