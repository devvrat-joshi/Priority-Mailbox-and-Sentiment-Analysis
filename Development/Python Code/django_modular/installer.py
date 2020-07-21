from shutil import copytree
import os
cd = os.getcwd()
if(os.path.exists("C:/Program Files/ModularAddinDjango")):
    print("already installed")
else: 
    copytree(cd+"\django_modular","C:/Program Files/ModularAddinDjango")
os.startfile(cd+"\django_modular\django_modular\publish\setup.exe")