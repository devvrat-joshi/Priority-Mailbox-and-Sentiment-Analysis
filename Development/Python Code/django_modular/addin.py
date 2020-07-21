import os
pathfile = open("C:/Program Files/ModularAddinDjango/pathdjango.txt",'r')
cd = pathfile.read()
pathfile.close()
os.system("cd {} & python threading_all.py".format(cd))