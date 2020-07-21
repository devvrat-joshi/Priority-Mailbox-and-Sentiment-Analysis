import senti
import check_outlook
import watch_maildata
import run
import time
from multiprocessing import Process
if __name__ == "__main__":
    print("Initializing Models")
    time.sleep(1)
    print("Wait for 40 seconds to complete the initialization")
    time.sleep(1)
    print("Allocating Resources")
    print("Core 1: Sentiment Analyzer")
    print("Core 2: Outlook On Off noter")
    print("Core 3: Outlook User Action Events Watcher")
    print("Core 4: Loading Database, Running Django")
    Process(target=senti.main).start()
    Process(target=check_outlook.main).start()
    Process(target=watch_maildata.main).start()
    Process(target=run.main).start()