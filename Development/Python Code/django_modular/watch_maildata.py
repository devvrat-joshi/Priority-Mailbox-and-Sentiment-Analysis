import os, time
def main():
    path_to_watch = "maildata/"
    before = dict ([(f, None) for f in os.listdir (path_to_watch)])
    while 1:
        time.sleep (10)
        after = dict ([(f, None) for f in os.listdir (path_to_watch)])
        added = [f for f in after if not f in before]
        removed = [f for f in before if not f in after]
        if added: 
            file = open("watch.txt",'w')
            file.write(str(added[0]))
            file.close()
        # if removed: 
        #     print ("Removed: ", ", ".join (removed))
        before = after