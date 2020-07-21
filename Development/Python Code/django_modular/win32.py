import win32com.client
import win32com
import os
import sys
from datetime import datetime
# import antispam
# from textblob import TextBlob
print("GO")

class mailfromost:
    From : str
    date: str
    body: str
    subject: str
    sentiment: str
    bodyc: str
    From1: str
    body1: str
    subject1: str

def emailleri_al(folder):
    messages = folder.Items
    print(messages)
    a=len(messages)
    listi = []
    current = datetime.now()
    current_date,current_month = current.day,current.month
    if a>0:
        i = 0
        for message2 in messages:
            i+=1
            print(message2.SenderName)
            if(i>8000):
                break
            # try:
            # if(int(str(message2.receivedtime)[8:10])>=current_date-10 and (int(str(message2.receivedtime)[5:7])<=current_month)):
                new = mailfromost()
                new.From = message2.SenderEmailAddress.replace(',','')
                new.bodyc = message2.body.replace(',',';').replace('\n','').replace('\r','')
                new.date = datetime.now()
                new.subject = message2.subject
                print(message2.entryid)
                date,month = int(str(message2.receivedtime)[8:10]),int(str(message2.receivedtime)[5:7])
                # if(int(str(message2.receivedtime)[8:10])>
                listi.append(new)
            # senti = TextBlob(new.bodyc).sentiment.polarity
            # print(j.From,j.body,j.subject)
            # except:
            #     print("Error")
            #     pass

    listi.reverse()
    return listi




def fetching():
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    overall = []
    print(1)
    for account in [accounts[0]]:
        global inbox
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        folders = inbox.Folders
        pp = 0
        for folder in folders:
            pp+=1
            if(str(folder) == "Sent Items"):
                overall_send = emailleri_al(folder)
            if(str(folder) == "Deleted Items"):
                overall_delete = emailleri_al(folder)
    return overall_delete,overall_send
            # if(pp==4):
            #     return overall
print(fetching()[0])
# for i in fetching():
    # print(i.From,i.bodyc,i.subject)
# f = open('data.csv','w')
# r = dict()
# p = dict()
# for i in h:
#     try:
#         y = antispam.score(i.bodyc)
#         if(i.From not in r):
#             p[i.From] = 1
#             r[i.From] = y
#         else:
#             p[i.From]+=1
#             r[i.From] += y
#     except:
#         pass
# x = open('gotonetwork.csv','w')
# for i in h:
#     try:
#         y = antispam.score(i.bodyc)
#         f.write("{},{},{},{},{}\n".format(str(i.From),str(i.subject),str(i.bodyc),str(y),str(r[i.From]/p[i.From])))
#         x.write("{},{}\n".format(str(r[i.From]/p[i.From]),str(y)))
#     except:
#         pass
# x.close()
# f.close()

# x = open('senders.csv','w')

# for i in list(r.keys()):
#     x.write("{},{}\n".format(i,str(r[i]/p[i])))
# x.close()