# section IMPORTS
from django.shortcuts import render
from threading import Thread
import win32com.client
import win32com
import os
import sys
from datetime import datetime
# Create your views here.
def main(request):
    return render(request, 'index.html')
def sentiment(request):
    file = open('comp.txt','r')
    x=file.readline()
    p = x.split(',')
    neg = '{},{}'.format(p[4],p[5])
    neu = '{},{}'.format(p[2],p[3])
    pos = '{},{}'.format(p[0],p[1])
    num = '{}'.format(p[6])
    dci = {'neg':neg,'pos':pos,'neu':neu,'num':num}
    return render(request, 'sentiment.html',dci)
def toexcel(request):
    return render(request, 'addintoexcel.html')
def priority(request):
    return render(request, 'priority.html')
def excel(request):
    return render(request, 'excel.html')
def develop(request):
    return render(request, 'developers.html')
def custom(request):
    return render(request, 'customize.html')
def updatesentiment(request):
    neg = request.POST['neg']
    pos = request.POST['pos']
    neu = request.POST['neu']
    num = request.POST['num']
    dci = {'neg':neg,'pos':pos,'neu':neu,'num':num}
    neg = neg.split(',')
    pos = pos.split(',')
    neu = neu.split(',')
    pos = pos+neu+neg+[num]
    print(pos)
    file = open('comp.txt','w')
    file.write(",".join(pos)+",{}".format(str(num)))
    file.close()
    return render(request, 'sentiment.html',dci)


import sys
filepath = open("C:/Program Files/ModularAddinDjango/pathapp.txt",'r')
pathapp = filepath.read()
filepath.close()
sys.path.insert(1, pathapp)
from app.views import db_update
Thread(target=db_update).start()