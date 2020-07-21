from django.shortcuts import render
from threading import Thread
import win32com.client
import win32com
import os,math
import pythoncom
import sys
from datetime import datetime
from app.models import senders, email
import antispam
from operator import itemgetter
# Create your views here.

def emailleri_al(folder,type_of):
    messages = folder.Items
    print(messages)
    a=len(messages)
    listi = []
    # current = datetime.datetime.month()
    # current_date,current_month = current.day,current.month
    if a>0:
        i = 0
        for message2 in messages:
            i+=1
            if(i>8000):
                break

            try:
                # if(int(str(message2.receivedtime)[8:10])>=current_date-5 and (int(str(message2.receivedtime)[5:7])<=current_month)):
                new = email()
                new.id_mail = message2.PropertyAccessor.BinaryToString(message2.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102"))
                already = email.objects.filter(id_mail = message2.PropertyAccessor.BinaryToString(message2.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x300B0102")))
                if(len(already)==1):
                    if((already[0].completeflag==1 and already[0].isreplied_count==1) or type_of != "received"):
                        continue
                    if(already[0].completeflag==0):
                        mysender = senders.objects.filter(address = message2.SenderEmailAddress)
                        mysender.update(sender_opened_count = mysender[0].sender_opened_count + (1-message2.UnRead))
                        already.update(completeflag = 1)
                    if(already[0].isreplied_count==0):
                        mysender = senders.objects.filter(address = message2.SenderEmailAddress)
                        if(message2.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")>0):
                            mysender.update(sender_reply_count = mysender[0].sender_reply_count + 1)
                            already.update(isreplied_count = 1)
                            print(already[0].isreplied_count)
                    continue
                        
                if(type_of=="send"):
                    new.sender_address = message2.To
                else:
                    new.sender_address = message2.SenderEmailAddress
                    sendernew = senders.objects.filter(address = message2.SenderEmailAddress)
                    if len(sendernew) == 0:
                        newsender = senders()
                        newsender.name = message2.SenderName
                        newsender.address = message2.SenderEmailAddress
                        newsender.sender_reply_count = 0
                        newsender.sender_total_count = 1
                        newsender.sender_importance = 2
                        newsender.sender_opened_count = (1-message2.UnRead)
                        if(1-message2.UnRead):
                            new.completeflag = 1
                        newsender.save()
                        if(type_of=="received"):
                            if(message2.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")>0):
                                sendernew.sender_reply_count = 1
                                message2.isreplied_count = 1
                                print("isreplied_count")
                    else:
                        sender = sendernew[0]
                        if(type_of=="received"):
                            if(message2.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x10810003")>0):
                                sendernew.update(sender_reply_count = sender.sender_reply_count+1)
                                new.isreplied_count = 1
                                print("isreplied_count")
                            sendernew.update(sender_total_count = sender.sender_total_count+1,sender_opened_count = sender.sender_opened_count+(1-message2.UnRead))
                            if(1-message2.UnRead):
                                new.completeflag = 1

                new.body = message2.body.replace("\n","")
                new.date = message2.receivedtime
                new.subject = message2.subject
                new.type_of = type_of
                new.read= (1-message2.UnRead)
                new.sender_name = message2.SenderName
                new.received_time = message2.ReceivedTime
                new.save()
                print(message2.entryid)
            except:
                pass
    listi.reverse()
    return listi



def replied_till_addin(request):
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    x1_id = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, outlook)
    accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts
    for account in [accounts[0]]:
        global inbox
        inbox = outlook.Folders(account.DeliveryStore.DisplayName)
        folder = inbox.Folders
        folder = inbox.Folders["Sent Items"]
        emailleri_al(folder,"send")
        folder = inbox.Folders["Deleted Items"]
        emailleri_al(folder,"deleted")
        folder = inbox.Folders["Inbox"].Folders["Positive"]
        emailleri_al(folder,"received")
        folder = inbox.Folders["Inbox"].Folders["Neutral"]
        emailleri_al(folder,"received")
        folder = inbox.Folders["Inbox"].Folders["Negative"]
        emailleri_al(folder,"received")
        folder = inbox.Folders["Inbox"]
        emailleri_al(folder,"received")
    return render(request,"index.html")

def rating_updater(request):
    emails = email.objects.filter(type_of = "received")
    for emailx in emails:
        if(emailx.isreplied_count!=0 and emailx.completeflag == False):
            mysender = senders.objects.filter(address = emailx.sender_address)[0]
            senders.objects.filter(address = emailx.sender_address).update(sender_reply_count = mysender.sender_reply_count + emailx.isreplied_count)
            print("isreplied_count")
            email.objects.filter(id_mail = emailx.id_mail).update(completeflag = True)
    return render(request,"index.html")

def sender_importance(request):
    sender_all = senders.objects.all()
    for i in sender_all:
        senders.objects.filter(address = i.address).update(sender_importance = (pow(2,(i.sender_reply_count))*(i.sender_opened_count)))
    return render(request,"index.html")

def sender_rank(request):
    sender_all = senders.objects.all()
    ranklist = []
    for i in sender_all:
        ranklist.append((i.sender_importance,i.address))
    ranklist.sort(key = itemgetter(0),reverse=True)
    file = open("ranklist.txt","w")
    for i in ranklist:
        file.write(i[1]+"\n")
    file.close()
    grp = math.ceil(len(ranklist)/6)
    lenth = len(ranklist)
    i = 0
    file = open("maildata/red.txt","w")
    for i in range(0,grp):
        if(i>=lenth):
            break
        if(ranklist[i][1]==""):
            continue                
        file.write(ranklist[i][1]+'\n')
    file.close()
    file = open("maildata/orange.txt","w")
    for i in range(grp,2*grp):
        if(i>=lenth):
            break
        if(ranklist[i][1]==""):
            continue                
        file.write(ranklist[i][1]+'\n')
    file.close()
    file = open("maildata/yellow.txt","w")
    for i in range(2*grp,3*grp):
        if(i>=lenth):
            break
        if(ranklist[i][1]==""):
            continue                
        file.write(ranklist[i][1]+'\n')
    file.close()
    file = open("maildata/green.txt","w")
    for i in range(3*grp,4*grp):
        if(i>=lenth):
            break
        if(ranklist[i][1]==""):
            continue        
        file.write(ranklist[i][1]+'\n')
    file.close()
    file = open("maildata/blue.txt","w")
    for i in range(4*grp,5*grp):
        if(i>=lenth):
            break
        if(ranklist[i][1]==""):
            continue                
        file.write(ranklist[i][1]+'\n')
    file.close()
    file = open("maildata/purple.txt","w")
    for i in range(5*grp,6*grp):
        if(i>=lenth):
            break
        if(ranklist[i][1]==""):
            continue
        file.write(ranklist[i][1]+'\n')
    file.close()
    lastmodeltime = datetime.now()
    file = open("lastupdatetime.txt","w")
    file.write(str(lastmodeltime))
    file.close()
    return render(request,"index.html")

def reply_update(request):
    repliedmail = email.objects.filter(subject__startswith = "RE:",type_of = "send")
    for mails in repliedmail:
        if mails.completeflag==False:
            ind = mails.body.find("Subject:")
            BODYtoSearch = mails.body[ind:]
            ind = BODYtoSearch.find("\r")
            BODYtoSearch = BODYtoSearch[ind+5:]
            BODYtoSearch = BODYtoSearch[:8]
            finded = email.objects.filter(type_of = "received",body__startswith = BODYtoSearch)
            finded.update(first_reply_time = mails.received_time, isreplied_count=1)
    return render(request,"index.html")

def db_update():
    watch = open('watch.txt','r')
    prev = watch.readline()
    print("OK")
    filepath = open("C:/Program Files/ModularAddinDjango/pathopenoutlook.txt",'r')
    pathdb = filepath.read()
    filepath.close()
    djangopath = open("C:/Program Files/ModularAddinDjango/pathdjango.txt",'r')
    pathdjango = djangopath.read()
    djangopath.close()
    while(1):
        watch = open('watch.txt','r')
        new = watch.readline()
        if(prev!=new and new!=''):
            print(new)
            string  = pathdjango + "maildata/" + new
            file = open(string,'r')
            data5 = file.read()
            print(data5)
            date_time = data5
            print(date_time)
            file = open(pathdb,'r')
            a = email.objects.filter(id_mail=new[:-4]).update(mail_open_time = date_time,app_open_time = file.read(), read = True)
            file.close()
        prev = new

Thread(target=db_update).start()