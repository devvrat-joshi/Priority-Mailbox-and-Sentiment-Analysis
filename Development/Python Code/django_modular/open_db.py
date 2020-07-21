import os
path1 = open("c:/Program Files/ModularAddinDjango/pathdjango.txt",'r')
path = path1.read()
path1.close()
os.system("cd {} & python manage.py shell".format(path))