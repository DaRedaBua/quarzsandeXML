from log import *
import Bimail
from colorama import init, Fore, Style
import os
import datetime
import shutil
import re

# TODO: Summe mit minusfahrten Franz schicken...

# debugging/logging
stacktrace = ""
recipients = []

# input Data
edvbookPath = ""
saveintern = ""

# XML
XMLerrOutPath = ""
XMLoutPath = ""

# UserDefinedVariables
dz = 0          #Dieselzuschlag
curMonth = 0    #Leistungsmonat
curYear = 0     #Leistungsjahr
abrFilePath = ""
savespotPath = ""
extbookPath = ""

# ProgrammAblaufVariablen
#   KZOrders['LL123AB']
#       [
#           dataDict
#               ['geraet]: F-3Achser
#               ['LFS-Datum']: 03.05.2021
#               ...,
#           dataDict,
#           dataDict,...
#      ]
#   KZOrders['LL123AB'][0]['geraet'] = F-3Achser
KZOrders = {}
#   fehler-entry: [row,col,code] 0: Ausgelassene Int Fahrt, 1: zwei KZ, 2: Ausgelassene subfahrt 3: Minusgeschäft
fehler = []

notCalculated = 0
allIntDZ = 0
header = [""]*14


def main():
    loadLogs("logs.csv")
    loadConfig()
    askUser()


def loadConfig():
    log(2000, "loadConfig", "")

    global recipients
    global outFolder
    global errFolder

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


def createFolders():
    log(4000, "createFolders()", "")
    global curYear
    global curMonth

    indir = saveintern + "20" + curYear + "_" + curMonth
    exdir = saveintern + "Subfrächter_20" + curYear + "_" + curMonth

    log(4005, "Neuer Interner ouput-Ordner, indir: ", indir)
    log(4010, "Neuer Subfrächter Ausgabe-Ordner: ", exdir)

    if os.path.exists(exdir):
        shutil.rmtree(exdir)

    if os.path.exists(indir):
        shutil.rmtree(indir)

    os.mkdir(exdir)
    os.mkdir(indir)


def extractLicensePlates(orplate):
    log(5000, "extractLicensePlate(orplate); orplate: ", orplate)

    entr = orplate.replace('-','')
    entr = entr.replace('_4A','')

    if orplate == '':
        entr = "KEIN"

    log(5005, "entr für verarbeitung: ", entr)

    notpattern = re.compile('[a-z]')
    findpattern = re.compile('([A-Z]|[0-9]){3,9}')

    platesFound = []

    # TODO: besser machen - derzeit könnte ABCD ein kennzeichen sein aber ll12md nicht
    # TODO: trim
    # TODO: edgeCases ausprobieren...

    plates = entr.split('+')       # Zuerst wird am + gepslittet
    for plate in plates:
        parts = plate.split(',')    # Dann an einem Beistrich - muss man hintereinander machen
        for part in parts:
            nono = notpattern.search(part)  # Ausschlussverfahren - wenn nur buchstaben => ungültig
            if nono is None:
                yesyes = findpattern.search(part)   # Suchverfahren - Muss Großbuchstaben
                if yesyes is not None and len(part) < 10:
                    platesFound.append(part)

    log(5030, "platesFound: ", platesFound)

    return platesFound


# TODO: STEHEN LASSEN - WENN SICH EXCEL-LISTE VERSCHIEBT HIER ÄNDERN!
# readAbrSheet teilt Zeilen nach Kennzeichen auf.
# Als Ergebnis dieser Methode ist das Dictionary KZOrders befüllt. Pro dict-key ein array mit dict pro fahrt.
# KZOrders['LL123AB']
#   [
#       dataDict (data ist nur hier der variablen-Name für das Dict)
#           ['geraet]:F-3Achser
#           ['LFS-Datum']:F-Achser
#           ...,
#       dataDict,
#       dataDict,...
#   ]
# Fehler werden nicht als KZ-Orders-Einträge behandelt sondern als hinweise auf originale Zelle in grüner Excel-Liste
#   (Zeile, Spalte, Fehlercode 0/1/2/3)
def readAbrSheet(srcsheet):
    log(6000, "readAbrSheet(srcsheet); srcsheet: ", srcsheet)

    global notCalculated
    global allIntDZ
    global header
    global fehler
    global KZOrders

    # Sammelt Überschriften aus Zeile 10 und DANCH Zeile 9 - hat nix mit einem Auftrag speziell zu tun
    # Notwendig, weil Gerät, LFS-nr, kunde, etc. in Zeile 10 in B-K steht (siehe obere For-Loop)
    #    Menge, Stunden, Ger.Kosten, Mautkosten aber in Zeile 8 L-O
    #    Untere Loop überschreibt leere Zellen 10 L-O mit gefüllten Zellen 9 L-O
    for c in range(1, srcsheet.ncols):
        header[c-1] = srcsheet.cell(9, c).value
    for c in range(srcsheet.ncols-4, srcsheet.ncols):
        header[c-1] = srcsheet.cell(8, c).value

    log(6005, "header: ", header)

    # Sortiert Zeile für Zeile bei Kennzeichen ein
    for i in range(10, srcsheet.nrows):
        success = True

        # In data wird eine Zeile/Fahrt als Dict gespeichert.
        data = dict()
        data['geraet'] = srcsheet.cell(i, 1).value
        data['lfs_datum'] = srcsheet.cell(i, 2).value
        data['lfs_nr'] = srcsheet.cell(i, 3).value
        data['art_lfrnt'] = srcsheet.cell(i, 4).value
        data['art'] = srcsheet.cell(i, 5).value
        data['kunden'] = srcsheet.cell(i, 6).value
        data['baustelle'] = srcsheet.cell(i, 7).value
        data['zone'] = srcsheet.cell(i, 9).value
        data['einheit'] = srcsheet.cell(i, 10).value
        data['menge'] = srcsheet.cell(i, 11).value
        data['stunden'] = srcsheet.cell(i, 12).value
        data['ger_kosten'] = srcsheet.cell(i, 13).value
        data['mautk'] = srcsheet.cell(i, 14).value
        data['anmerkungen'] = ""
        data['zeile'] = i+1

        # Zone (alphanumerisch) fix in UPPER String umwandeln
        if isinstance(data['zone'], str):
            data['zone'] = data['zone'].upper()
        else:
            data['zone'] = str(int(data['zone']))

        # doFehler in verbindung mit Success:
        # Weiter unten überprüft mein Programm die einzelnen Zeilen auf verarbeitbarkeit
        # Wenn Fehler passieren, wird die glob Variable fehler appendet.
        # Allerdings sollen die Zeilen Treibstoff und die Summenzeile das nicht auslösen.
        # Quickfix: naturally true until proven false, siehe nächste Zwei IFs
        doFehler = True

        if "Treibstoff" in kz:
            allIntDZ = data['ger_kosten']
            doFehler = False

        if "Summe" in data['geraet']:
            doFehler = False

        # Herausfinden der Kennzeichen in dieser Zeile, gibt array zurück, [plates]
        kz = srcsheet.cell(i, 8).value
        plates = extractLicensePlates(kz)

        # Wenn mehr als 1 Kennzeichen in Zeile -> Leichter Error (wird im moment nicht behandelt)
        if(len(plates) > 1) and doFehler:
            fehler.append([i, 7, 1])

        # Kein Kennzeichen => Schwerer Fehler
        if(len(plates) == 0) and doFehler:
            success = False
            fehler.append([i, 7, 0])

        # Keine Gerätekosten
        if data['ger_kosten'] is None or data['ger_kosten'] == '' or data['ger_kosten'] == 0 and doFehler:
            fehler.append([i, 12, 0])
            success = False

        # Keine Lieferscheinnummer
        if data['lfs_nr'] is None or data['lfs_nr'] == '' or data['lfs_nr'] == 0:
            fehler.append([i, 2, 0])
            success = False

        # Wenn Treibstoff oder Summenzeile => nicht berechnen (suc false), aber keinen Fehler schreiben
        if not doFehler:
            success = False

        # Wenn Zeile soweit ok, Fahrt im Kennzeichen-Dictionary einordnen
        if success:
            if plates[0] not in KZOrders:   # Wenn Kennzeichen noch nicht im Dictionary => Anlegen
                KZOrders[plates[0]] = []
            data['kz'] = plates[0]          # Erst hier, weil certain len(plates) > 0
            KZOrders[plates[0]].append(data)    # data-Obj wird hinzugefügt
            log(6050, "Kennzeichen: ", data[kz])
            log(6051, "data: ", data)
        elif doFehler:
            notCalculated += 1              # Summe an ausgelassenen Datensätzen wird um 1 erhöht

main()