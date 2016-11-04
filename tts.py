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
from pygame import mixer
import re


class Countmoney(object):

    def __init__(self, _path):
        self.speaker        = win32com.client.Dispatch("SAPI.SpVoice")
        self.path           = getfolder.get_path(getfolder.FOLDERID.SavedGames, getfolder.UserHandle.current) + "\\Frontier Developments\\Elite Dangerous" if not _path else _path
        self.lastEvent      = (datetime.datetime.utcnow() - datetime.timedelta(1)).replace(tzinfo=pytz.utc)
        self.bounty         = 0
        self.bountyCount    = 1       
        self.czbounty       = 0
        self.czCount        = 1
        self.speakEvery     = 1000000 #Speak about money every this many credits
        self.SVSFIsXML      = 8 #This flag is the SpeechVoiceSpeakFlag that controls whether the text is read as is or is interpreted as XML. 
                                #https://msdn.microsoft.com/en-us/library/ms720892(v=vs.85).aspx further reading
        mixer.init()
        mixer.music.load('cena.mp3')

        pass

    def say(self, stringing, isLiteral=False):
        string          = ""
        
        if isLiteral:
            string = re.sub('(\d+)', '<spell>' + r'\1' + '</spell>', stringing)              
            string = re.sub('(-)', ' dash ', string)                                    #don't do it this way ever
            string = re.sub('([A-Z][A-Z]+)', '<spell>' + r'\1' + '</spell>', string)    #ugh
            self.speaker.Speak(string, self.SVSFIsXML)

        else:
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
                line = fh.readline()
                if not line:
                    time.sleep(1)
                    continue

                parsedJSON      = json.loads(line)
                timestamp       = iso8601.parse_date(parsedJSON['timestamp'])

                if timestamp >= self.lastEvent:
                    self.lastEvent      = timestamp
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
            elif(string['event'] == 'FactionKillBond'):
                self.combatBonds(string)

            #Jumping
            elif(string['event'] == 'FSDJump'):
                self.Jump(string)
        
        except KeyError:
            return
        
        return

    def Jump(self, string):
        self.say('Arrived at ' + string['StarSystem'], True)

    def bounties(self, string):
        self.bounty     += string['Rewards'][0]['Reward']
        if(not mixer.music.get_busy()):
            #mixer.music.play()
            pass

        if self.bountyCount * self.speakEvery <= self.bounty:
            self.say('You have accumulated ' + str(self.bountyCount * self.speakEvery) + ' in bounties.', False)
            self.bountyCount += 1

        return


    def combatBonds(self, string):
        self.czbounty     += string['Reward']
        if(not mixer.music.get_busy()):
            #mixer.music.play()
            pass

        if self.czCount * self.speakEvery <= self.czbounty:
            self.say('You have accumulated over ' + str(self.czCount * self.speakEvery) + ' in combat bonds.', False)
            self.czCount += 1

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