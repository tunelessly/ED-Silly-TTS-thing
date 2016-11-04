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
from pygame import mixer
import re
import tkinter as tk
import tkinter.messagebox
import threading
import zlib
import base64




class JournalTTS(tk.Frame):

    def __init__(self, _path, master = None):
        super().__init__(master)
        self.speaker            = win32com.client.Dispatch("SAPI.SpVoice")
        self.path               = getfolder.get_path(getfolder.FOLDERID.SavedGames, getfolder.UserHandle.current) + "\\Frontier Developments\\Elite Dangerous" if not _path else _path
        self.lastEvent          = (datetime.datetime.utcnow() - datetime.timedelta(1)).replace(tzinfo=pytz.utc)
        self.bounty             = 0
        self.bountyCount        = 1       
        self.czbounty           = 0
        self.czCount            = 1
        self.speakEvery         = 1000000 #Speak about money every this many credits
        self.SVSFIsXML          = 8 #This flag is the SpeechVoiceSpeakFlag that controls whether the text is read as is or is interpreted as XML. 
                                    #https://msdn.microsoft.com/en-us/library/ms720892(v=vs.85).aspx further reading
        self.thread             = None
        self.threadShouldStop   = False
        mixer.init()
        
        try:
            mixer.music.load('cena.mp3')
        except:
            pass

        self.master.title("Silly ED TTS Thing")
        self.createWidgets()
        self.master.protocol('WM_DELETE_WINDOW', self.onQuitButtonPressed)
        self.master.minsize(width=250, height=100)
        self.master.resizable(width=False, height=False)
        pass

    def reset(self):
        self.lastEvent          = (datetime.datetime.utcnow() - datetime.timedelta(1)).replace(tzinfo=pytz.utc)
        self.bounty             = 0
        self.bountyCount        = 1       
        self.czbounty           = 0
        self.czCount            = 1

    def createWidgets(self):
        self.bountyCheckboxVar  = tk.BooleanVar()
        self.czCheckboxVar      = tk.BooleanVar()
        self.jumpCheckboxVar    = tk.BooleanVar()

        self.bountyCheckbox     = tk.Checkbutton(self.master, text="Read out bounties", variable=self.bountyCheckboxVar, command=self.bountyCheckboxStateChanged)
        self.czCheckbox         = tk.Checkbutton(self.master, text="Read out combat bonds", variable=self.czCheckboxVar, command=self.czCheckboxStateChanged)
        self.jumpCheckbox       = tk.Checkbutton(self.master, text="Read out jumps", variable=self.jumpCheckboxVar, command=self.jumpCheckboxStateChanged)
        self.startButton        = tk.Button(self.master, text="Start", command=self.startWatching)

        self.bountyCheckbox.pack()
        self.czCheckbox.pack()
        self.jumpCheckbox.pack()
        self.startButton.pack()
        
        pass

    def startWatching(self):
        #dirty reads, dirty reads everywhere!

        self.startButton['text'] = 'Restart'
        if not self.thread:
            self.thread             = threading.Thread(target=self.watchFile)
            self.thread.daemon      = True
            self.thread.start()
        else:
            self.reset()
            self.threadShouldStop   = True
            self.thread.join()
            self.threadShouldStop   = False
            self.thread             = threading.Thread(target=self.watchFile)
            self.thread.daemon      = True
            self.thread.start()


    def bountyCheckboxStateChanged(self):
        #print(self.bountyCheckboxVar.get())
        pass

    def czCheckboxStateChanged(self):
        #print(self.czCheckboxVar.get())
        pass

    def jumpCheckboxStateChanged(self):
        #print(self.jumpCheckboxVar.get())
        pass

    def onQuitButtonPressed(self):
        if self.thread:
            self.threadShouldStop   = True
            self.thread.join()
        
        self.master.destroy()

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
            tk.messagebox.showwarning("Open File", "Path not found: " + self.path)
            return
        except ValueError:
            tk.messagebox.showwarning("Open File", "No log files found at: " + self.path)
            return
        

        try:
            while True:
                if(self.threadShouldStop):
                    fh.close()
                    return

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
            return
        except json.decoder.JSONDecodeError:
            tk.messagebox.showwarning("Invalid JSON", "Invalid JSON. Maybe it's not a valid Elite:Dangerous Journal file?")
            return
        
        return



    def parseEvents(self, string):
        try:
            #Normal Bounties
            if(string['event'] == 'Bounty' and self.bountyCheckboxVar.get()):
                self.bounties(string)

            #CZ Combat bonds
            elif(string['event'] == 'FactionKillBond' and self.czCheckboxVar.get()):
                self.combatBonds(string)

            #Jumping
            elif(string['event'] == 'FSDJump' and self.jumpCheckboxVar.get()):
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
    rootWindow      = tk.Tk()
    try:
        thing = JournalTTS(argv[1], master=rootWindow)
    except IndexError:
        thing = JournalTTS(None, master=rootWindow)
    
    #thing.watchFile()

    thing.mainloop()
    return 0

if __name__ == '__main__':
    main(sys.argv)