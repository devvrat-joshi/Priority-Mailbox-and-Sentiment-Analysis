from textblob import TextBlob
import time
print("all ok")
def main():
    filepath = open("C:/Program Files/ModularAddinDjango/pathtest.txt",'r')
    pathtest = filepath.read()
    filepath.close()
    filepath = open("C:/Program Files/ModularAddinDjango/pathpol.txt",'r')
    pathpol = filepath.read()
    filepath.close()
    testfile = open(pathtest, 'r')
    prev = 50
    while(1):
        time.sleep(0.1)
        try:
            testfile = open(pathtest, 'r')
            curr = testfile.read()
            if(curr!=prev):
                f = open(pathpol, 'w')
                try:
                    x = TextBlob(curr)
                    pol = x.sentiment.polarity
                    print(pol)
                except:
                    pol = 0
                    pass
                comparatorfile = open("comp.txt",'r')
                comp = list(map(float,comparatorfile.readline().split(',')))
                if(pol>comp[0] and pol<=comp[1]):
                    f.write('1')
                elif(pol>=comp[2] and pol<=comp[3]):
                    f.write('2')
                elif(pol>=comp[4] and pol<comp[5]):
                    f.write('3')
                f.close()
                prev= curr
        except:
            pass
        time.sleep(0.3)
# So all mails moved to their respective folders, here 10 mails at a time are moved, But if some mails are not mailitems so they are not moved like system administration mails. Also we can change the number of mails moved at a time.