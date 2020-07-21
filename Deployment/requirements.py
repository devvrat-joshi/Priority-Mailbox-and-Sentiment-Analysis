import os
try:
    import django.contrib
except:
    print("installing requirement 1")
    os.system("pip install Django")
try:
    import textblob
except:
    print("installing requirement 2")
    os.system("pip install textblob")
try:
    import antispam
except:
    print("installing requirement 3")
    os.system("pip install antispam")
try:
    import win32com
except:
    print("installing requirement 4")
    os.system("pip install pywin32==224")
try:
    import antispam
except:
    print("installing requirement 5")
    os.system("pip install antispam")
os.system("pip install django-sslserver")