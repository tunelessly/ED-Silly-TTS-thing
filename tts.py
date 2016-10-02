import win32com.client
import glob
import os
import sys
import time
import json
import datetime
import pytz
import iso8601
from threading import Thread


class Countmoney(object):

    def __init__(self, _path):
        self.speaker    = win32com.client.Dispatch("SAPI.SpVoice")
        self.path       = "C:/Users/" + os.getlogin() + "/Saved Games/Frontier Developments/Elite Dangerous/" if not _path else _path
        self.lastEvent  = (datetime.datetime.utcnow() - datetime.timedelta(1)).replace(tzinfo=pytz.utc)
        self.bounty     = 0
        pass

    def say(self, stringing):
        self.speaker.Speak(stringing)
        return

    def watchFile(self):
        fp = None 
        fh = None
        try:
            fp         = max(glob.iglob(self.path + "\\*.log"), key=os.path.getctime)
            fh         = open(fp, mode='r')
        except FileNotFoundError:
            print ("Path not found.")
            return
        except ValueError:
            print("No log files found.")
            return
        
        print("Opened file: " + fp)

        try:
            while True:
                time.sleep(2)
                lines    = fh.readlines()

                for line in lines:
                    parsedJSON  = json.loads(line)
                    timestamp   = iso8601.parse_date(parsedJSON['timestamp'])
                    
                    if timestamp > self.lastEvent:
                        self.lastEvent = timestamp
                        self.parseEvents(parsedJSON)
        except KeyboardInterrupt:
            print("Quitting")
            return
        except json.decoder.JSONDecodeError:
            print("Invalid JSON. Maybe it's not a valid Elite:Dangerous Journal file?'")
            return
        
        pass



    def parseEvents(self, string):
        try:
            if(string['event'] == 'Bounty'):
                self.bounty     += string['Reward'] 

            if self.bounty - 1000000 >= 0:
                self.say('You have accumulated one million in bounties.')
                self.bounty = self.bounty - 1000000
        except KeyError:
            pass
        
        return
        
def main(argv):
    try:
        thing = Countmoney(argv[1])
    except IndexError:
        thing = Countmoney(None)
    
    thing.watchFile()
    return 0

if __name__ == '__main__':
    main(sys.argv)