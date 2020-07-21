import urllib.request,time
while(1):
    try:
        urllib.request.urlopen('http://127.0.0.1:8000/second.html')
    except:
        pass
    time.sleep(30)