from mailmerge import MailMerge
import datetime
import pandas as pd
import win32com.client

template = "C:\\Users\\braaten\\Desktop\\Letter of Request Template.docx"
excelfile = "C:\\Users\\braaten\\Desktop\\LOR Worksheet.xlsx"
df = pd.read_excel(excelfile, sheet_name = "Sheet1", usecols="A:Y").fillna('')
organization=df["Organization"].tolist()
firstname=df["First Name"].tolist()
lastname=df["Last Name"].tolist()
address=df["Address 1"].tolist()
city=df["City"].tolist()
state=df["State"].tolist()
zipcode=df["ZIP"].tolist()
amount=df["Amount"].tolist()
bullet1=df["Bullet 1"].tolist()
bullet2=df["Bullet 2"].tolist()
bullet3=df["Bullet 3"].tolist()
bullet4=df["Bullet 4"].tolist()
bullet5=df["Bullet 5"].tolist()
bullet6=df["Bullet 6"].tolist()
bullet7=df["Bullet 7"].tolist()
bullet8=df["Bullet 8"].tolist()
bullet9=df["Bullet 9"].tolist()
bullet10=df["Bullet 10"].tolist()
filename=df["Filename"].tolist()
email=df["Email"].tolist()
requested_sponsor_level=df["Requested Sponsor Level"].tolist()
requested_by_date=df["Requested by Date"].tolist()

word = win32com.client.Dispatch('Word.Application')
for i in range (0,len(organization)):
    document = MailMerge(template)
    document.merge(
        
        Organization=organization[i],
        First_Name=firstname[i],
        Last_Name=lastname[i],
        Address_1=address[i],
        City=city[i],
        State=state[i],
        ZIP=str(zipcode[i]),
        Amount="$"+format(amount[i],","),
        Email=str(email[i]),
        Requested_Sponsor_Level=str(requested_sponsor_level[i]),
        Requested_by_Date=f'{requested_by_date[i].to_pydatetime():%B} {requested_by_date[i].to_pydatetime().day}, {requested_by_date[i].to_pydatetime().year}',
        )
    for x in range (0,9):
        if not bullet1[i]:
            break
        else:
            document.merge(Bullet_1="●    "+bullet1[i])
        if not bullet2[i]:
            break
        else:
            document.merge(Bullet_2="●    "+bullet2[i])
        if not bullet3[i]:
            break
        else:
            document.merge(Bullet_3="●    "+bullet3[i])
        if not bullet4[i]:
            break
        else:
            document.merge(Bullet_4="●    "+bullet4[i])
        if not bullet5[i]:
            break
        else:
            document.merge(Bullet_5="●    "+bullet5[i])
        if not bullet6[i]:
            break
        else:
            document.merge(Bullet_6="●    "+bullet6[i])
        if not bullet7[i]:
            break
        else:
            document.merge(Bullet_7="●    "+bullet7[i])            
        if not bullet8[i]:
            break
        else:
            document.merge(Bullet_8="●    "+bullet8[i])              
        if not bullet9[i]:
            break
        else:
            document.merge(Bullet_9="●    "+bullet9[i])  
        if not bullet10[i]:
            break
        else:
            document.merge(Bullet_10="●    "+bullet10[i])  
            
    document.write("C:\\Users\\braaten\\Desktop\\LORs\\Letter of Request - "+organization[i]+" - "+str(amount[i])+".docx")
    document.close()
word.Quit()

word = win32com.client.Dispatch('Word.Application')
for i in range (0,len(organization)):
    in_file = "C:\\Users\\braaten\\Desktop\\LORs\\Letter of Request - "+organization[i]+" - "+str(amount[i])+".docx"
    out_file= "C:\\Users\\braaten\\Desktop\\LORs\\Letter of Request - "+organization[i]+" - "+str(amount[i])+".pdf"
    doc = word.Documents.Open(in_file)
    doc.SaveAs(out_file, FileFormat=17)
    doc.Close()
word.Quit()

