import os
def main():
    pathfile = open("C:/Program Files/ModularAddinDjango/pathdjango.txt",'r')
    pathdjango = pathfile.read()
    pathfile.close()
    os.system("cd {} & python manage.py runserver".format(pathdjango))
