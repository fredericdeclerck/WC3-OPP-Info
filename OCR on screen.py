# -*- coding: utf-8 -*-
"""
Created on Sun May  3 20:35:10 2020

@author: fred
"""
#This program is meant to leverage the data I am gathering over time
#from my opponents, in an excel file
#and let me know quickly about the strat(s) my opponent has done in the past
# provided i have played him before.

#To accomplish so this program does the following:
# 1 - captures the screen upon pressing F6
#     (user needs to press F6 while game loads)
# 2 - based on that capture, retrieve opp name, using OCR
# 3 - then look for a match in the excel file
#     if any info is available, it text-to-speeches
#     it out (as the game finishes to load or has just started)

###############################
#libraries
import pandas as pd
import pyscreenshot           #library for screenshots
import ocrspace               #library for OCR, it is online, i got a free license
from pynput import keyboard   #library for keyboard events
import pyttsx3                #library for text to speech

###############################
#Constants
OCR_API = ocrspace.API("d92462a6fd88957") #key to use API
MY_OWN_NAME = "FREDOKUN"                #my login
OPP_COL_TITLE = "Opponent"              #title containing the name
#File with meta data
OPP_FILE = 'M:/Installed Apps/blizz/Warcraft III/Opp-history.xlsx'
DATA_WORKSHEET = 'Data'                 #data worksheet
IM_BUFF = 'imBuff.png'                  #file generated with the screenshot
STR_DELIMITER = '\n'         #character in the returned string between player names
NOT_ABLE_TO_READ = "not able to read image"

###############################
#global counter
#because of keyboard events bug
pressAndReleaseCounter = 0

###############################
#Returns the text found on the image
def ReadImage(file):
    try:
        return OCR_API.ocr_file(file)
    except:
        return NOT_ABLE_TO_READ

###############################
#Opens up the excel file and gets the data
def loadOppsData():
    df = pd.read_excel(OPP_FILE, sheet_name=DATA_WORKSHEET)
    return df

###############################
#Captures the area of the screen where the names are
#saves it in a local file
#the names of both my opponent and I are there
#we dont know in what order though
#so the area needs to capture both locations
def CapturePlayersNames():
    #coordinates where teh names are located
    #while game is loading
    im = pyscreenshot.grab(bbox=(330, 500, 1600, 550))
    #im.show()
    im.save(IM_BUFF) #overwrites image

###############################
#from a string, retrieves a potential opponent name
#not mine  , 1 problem is that we dont know
#whose name comes 1st
#this function gets the 1st string, compares it to fredokun
#and gets the next one if needed
def retrieveOppName(myString):
    #players are separated by '\n'
    retVal = 1
    start = -1
    positions = []
    while retVal > 0: #needs to find the positions of the delimiters
        retVal = myString.find(STR_DELIMITER, start+1)
        start = retVal
        positions.append(retVal)
    #print( positions)

    #creates a string from the letters in between delimiters
    potentialCulprit = myString[0: positions[0]-1]
    if potentialCulprit == MY_OWN_NAME: #then look for someone else but me
        potentialCulprit = myString[int(positions[0]+1): int(positions[1]-1)]
    print(potentialCulprit)
    return potentialCulprit

###############################
#if player is found then return info,
#it works if the player has mutiple entries
#the excel file needs to have opp name in lower case
def findPlayerinFile(df, name):
     namesMatching = df[OPP_COL_TITLE].str.find(name.lower())
     myText = " "
     for i in range(len(df[OPP_COL_TITLE])):
         if namesMatching[i] == 0: #if = 0 it matches,  else, it is -1
             print(df.iloc[i])
             for j in range(len(df.iloc[i])):
                myText += str(df.iloc[i, j])
                myText += "," #commas slow down the text to speech
     if myText == " ":
        myText = "Player " + str(name) + " not found in database"

     return myText


####I wish this worked, it is more elegant than currenlty
# def on_key_press(key):
#     if key == keyboard.Key.f6:
#       CapturePlayersNames()
#       detectedText = ReadImage(IM_BUFF)
#      # print (detectedText)
#       if len (detectedText) > 0:
#           opponent = retrieveOppName(detectedText)
#           oppInfo = findPlayerinFile(OppsData, opponent)
#       else:
#           oppInfo = "Player not found"
#       engine.say( oppInfo )
#       engine.runAndWait()
#     elif key == keyboard.Key.f12:
#       return False
#     else:
#       #print('Received event {}'.format(event))
#       print(" _-_ ")

###############################
######             Main
OppsData = loadOppsData()   #excel data is loaded
engine = pyttsx3.init()     #starts the text to speech
#change the speech pace, assuming 100 is default
engine.setProperty('rate', 125)


###I wish this worked
# with keyboard.Listener(on_press = on_key_press) as listener:
#     listener.join()

# The event listener will be running in this block
#essentially looking for F6 to "work" or F12 to exit
with keyboard.Events() as events:
    for event in events:
        if event.key == keyboard.Key.f6:
            print("   F6  ")
            pressAndReleaseCounter += 1   ###need this as events come in 2..
            if (pressAndReleaseCounter % 2) == 0:##one press, one release
                CapturePlayersNames()
                detectedText = ReadImage(IM_BUFF)
                print (detectedText)
                if detectedText == NOT_ABLE_TO_READ:
                    oppInfo = "OCR API did not return data"
                elif len(detectedText) > 0:
                    opponent = retrieveOppName(detectedText)
                    if len(opponent) > 0:
                        oppInfo = findPlayerinFile(OppsData, opponent)
                    else:
                        oppInfo = "OCR unable to detect name"
                else:
                    oppInfo = "OCR unable to detect name"

                engine.say(oppInfo)
                engine.runAndWait()
        elif event.key == keyboard.Key.f12:
            break
       # else:
       #     #print('Received event {}'.format(event))
       #     print(event.key)
