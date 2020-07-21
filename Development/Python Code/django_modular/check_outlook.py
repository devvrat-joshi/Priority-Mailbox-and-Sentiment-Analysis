import subprocess
def process_exists(process_name):
    call = 'TASKLIST', '/FI', 'imagename eq %s' % process_name
    # use buildin check_output right away
    output = subprocess.check_output(call).decode()
    # check in last line for process name
    last_line = output.strip().split('\r\n')[-1]
    # because Fail message could be translated
    return last_line.lower().startswith(process_name.lower())
import time,datetime
def main():
    while(1):
        count = 0
        while(process_exists("outlook.exe")):
            count+=1
            if(count==1):
                file = open("outlook_open.txt","w")
                file.write(str(datetime.datetime.now()))
                file.close()
                print("Noted")