import win32com.client
import win32com
import os
import sys
from datetime import datetime
import antispam
from textblob import TextBlob
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
    if a>0:
        i = 0
        for message2 in messages:
            i+=1
            if(i>8000):
                break
            try:
                print(i)
                new = mailfromost()
                new.From = message2.SenderEmailAddress.replace(',','')
                new.bodyc = message2.body.replace(',',';').replace('\n','').replace('\r','')
                new.date = datetime.now()
                senti = TextBlob(new.bodyc).sentiment.polarity
                if(senti<0.05 and senti>-0.05):
                    new.sentiment = 'Neutral \N{neutral face}'
                elif(senti>0.05 and senti<0.7):
                    new.sentiment = 'Positive \U0001F642'
                elif(senti>-0.7 and senti<-0.05):
                    new.sentiment = 'Negative \U0001F641'
                elif(senti>0.7):
                    new.sentiment = 'Extreme Positive \U0001F60D'
                elif(senti<-0.7):
                    new.sentiment = 'Extreme Negative \U0001F62D'
                new.body = new.bodyc[:100]
                new.subject = message2.subject.replace(',',';')
                listi.append(new)
                #print(j.From,j.body,j.subject)
            except:
                print("Error")
                #print(account.DeliveryStore.DisplayName)
                pass

            #  try:
            #     message2.Save
            #     message2.Close(0)
            #  except:
            #      pass
    listi.reverse()
    return listi



def fetching():
    
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    overall = []
    print(1)
    for account in [accounts[1]]:
        global inbox
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        folders = inbox.Folders
        pp = 0
        for folder in folders:
            pp+=1
            overall = emailleri_al(folder)
            print(overall)
            if(pp==1):
                return overall
                break
h = fetching()
f = open('data.csv','w')
r = dict()
p = dict()
for i in h:
    try:
        y = antispam.score(i.bodyc)
        if(i.From not in r):
            p[i.From] = 1
            r[i.From] = y
        else:
            p[i.From]+=1
            r[i.From] += y
    except:
        pass
x = open('gotonetwork.csv','w')
for i in h:
    try:
        y = antispam.score(i.bodyc)
        f.write("{},{},{},{},{}\n".format(str(i.From),str(i.subject),str(i.bodyc),str(y),str(r[i.From]/p[i.From])))
        x.write("{},{}\n".format(str(r[i.From]/p[i.From]),str(y)))
    except:
        pass
x.close()
f.close()

x = open('senders.csv','w')

for i in list(r.keys()):
    x.write("{},{}\n".format(i,str(r[i]/p[i])))
x.close()