from colorama import init, Fore, Style

stackTrace = ""
allTrace = ""
allowedLogs = []

# log(123, "Log Event: ", var_of_interest)


def loadLogs(path):
    with open(path, 'r') as ctrl:
        lines = ctrl.readlines()
        for line in lines:
            x = line.replace("\n","")
            x = x.strip()
            allowedLogs.append(x)

    print(allowedLogs)


def log(nr, text, var):
    global stackTrace
    global allTrace

    if str(nr) in allowedLogs:
        print(Fore.WHITE + Style.NORMAL + str(nr) + str(text) + str(var))
        stackTrace = stackTrace + "\n" + str(nr) + str(text) + str(var)

    allTrace = allTrace + "\n" + str(nr) + str(text) + str(var)

def getStackTrace():
    return stackTrace

def getAllTrace():
    return allTrace
