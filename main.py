from log import *
import Bimail
from colorama import init, Fore, Style
import os

# debugging/logging
stacktrace = ""
recipients = []

#input Data
edvbookPath = ""
saveintern = ""
abrFilePath = ""
savespotPath = ""
extbookPath = ""

#XML
XMLerrOutPath = ""
XMLoutPath = ""

#WorkingVariables
dz = 0          #Dieselzuschlag
curMonth = 0    #Leistungsmonat
curYear = 0     #Leistungsjahr

def main():
    loadLogs("logs.csv")
    loadConfig()
    askUser()

def loadConfig():
    log(2000, "loadConfig", "")

    global recipients
    global inFolder
    global outFolder
    global errFolder
    global backupFolder
    global AgMappingPath
    global TourMappingPath
    global sourcePath

    global edvbookPath
    global saveintern
    global abrFilePath
    global savespotPath
    global extbookPath

    with open('config.csv', 'r') as f:
        lines = f.readlines()
        print(lines)
        for x in range(len(lines)):
            lineElems = lines[x].split(';')

            if x == 0:
                edvbookPath = lineElems[1]
                log(2005, "edvbookPath ", edvbookPath)

            if x == 1:
                saveintern = lineElems[1]
                log(2010, "savintern ", saveintern)

            if x == 2:
                abrFilePath = lineElems[1]
                log(2015, "abrFilePath ", abrFilePath)

            if x == 3:
                savespotPath = lineElems[1]
                log(2020, "savespotPath ", savespotPath)

            if x == 4:
                extbookPath = lineElems[1]
                log(2025, "extbookPath ", extbookPath)

            if x == 5:
                for e in range(1,len(lineElems)):
                    recipients.append(lineElems[e])
                log(2030, "recipients ", recipients)


def askUser():

    log(3000, "askUser()", "")

    global dz
    global curYear
    global curMonth

    print(Fore.CYAN + Style.BRIGHT + "Willkommen im Quarzsande Programm v2.0!")

    print(Fore.CYAN + Style.BRIGHT + "Geben Sie bitte den Treibstoffzuschlag in % ein! (Bei Subfrächtergutschriften wird automatisch +2% Aufschlag gerechnet!)")
    dp = input()
    dp = dp.replace("%", "")
    dp = dp.replace(",", ".")
    dp = float(dp)
    dp /= 100
    dp += 1
    dz = dp

    log(3005, "dz: ", dz)

    while curMonth <= 0 or curMonth > 12 or curYear < 21 or curYear > 50:
        print(Fore.CYAN + Style.BRIGHT + "\nGeben Sie den Leistungszeitraum ein! mm/jj - z.B.: '09/21'")
        inp = input()
        inp = inp.split("/")
        curMonth = int(inp[0])
        curYear = int(inp[1])
        log(3010, "curMonth: ", curMonth)
        log(3011, "curYear: ", curYear)


main()