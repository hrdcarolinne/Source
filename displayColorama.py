from colorama import *

def resetColor():
    print(Style.RESET_ALL, end='')
    print(Fore.BLUE + Back.LIGHTWHITE_EX + Style.NORMAL , end='')
    

def inputColor(prompt):
    print(Fore.YELLOW + prompt + Back.LIGHTYELLOW_EX, end='')
    userInput = input(Fore.BLACK + Back.LIGHTYELLOW_EX + Style.BRIGHT)

    resetColor()

    return userInput

def printError(prompt):
    print(Back.LIGHTRED_EX + Fore.RED + Style.DIM + prompt)

    resetColor()

def printColor(prompt, styleColorama, endWith='\n'):
    print(styleColorama + prompt, end=endWith)

    resetColor()

def displayDivision():
    print('\n' +Fore.CYAN + Back.LIGHTCYAN_EX + Style.DIM + '*************************************************************************************')
 
    resetColor()