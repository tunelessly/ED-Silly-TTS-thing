import win32com.client
import glob
import os
import sys
import time
import json
import datetime
import pytz
import iso8601
import getfolder
from threading import Thread


class Countmoney(object):

    def __init__(self, _path):
        self.speaker    = win32com.client.Dispatch("SAPI.SpVoice")
        self.path       = getfolder.get_path(getfolder.FOLDERID.SavedGames, getfolder.UserHandle.current) + "\\Frontier Developments\\Elite Dangerous" if not _path else _path
        self.lastEvent  = (datetime.datetime.utcnow() - datetime.timedelta(1)).replace(tzinfo=pytz.utc)
        self.bounty     = 0
        self.czbounty   = 0
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
        
        return



    def parseEvents(self, string):
        try:

            #Normal Bounties
            if(string['event'] == 'Bounty'):
                self.bounties(string)

            #CZ Combat bonds
            if(string['event'] == 'Bounty'):
                self.czbounty(string)
        
        except KeyError:
            return
        
        return

    def bounties(self, string):
        self.bounty     += string['Rewards'][0]['Reward']

        if self.bounty - 1000000 >= 0:
            self.say('You have accumulated one million in bounties.')
            self.bounty = self.bounty - 1000000

        return


    def combatBonds(self, string):
        self.czbounty     += string['Reward']

        if self.czbounty - 1000000 >= 0:
            self.say('You have accumulated one million in combat bonds.')
            self.czbounty = self.czbounty - 1000000

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