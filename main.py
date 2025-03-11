import os
from logData import *
from quantumConnection import *
from salesforceConnection import *
from workOrder import *
from validation import *
from standardPricing import *
from contractPricing import *
from agreementPricing import *
from displayColorama import *


def clear_screen():
    os.system('cls||clear')
        
    resetColor()

def testConnections():
    print('\n\tEstablishing connection...', end='\t')
    try:
        if not testQuantumConnection(): 
            return False
        if not testSalesforceConnection():
            return False
        if not testAccessConnection():
            return False
        print('\n')
        return True
    
    except Exception as e:
        return False
    
def startPricing():
    
    wo = None

    workOrderNumber = inputColor('\nPlease enter Work Order Number:\t').strip().upper()

    wo = WorkOrder(workOrderNumber)

    if wo._workOrderNumber == '': 
        displayDivision()
        startPricing()

        return 
    
    if not Validation(wo).checkForPMAs():
        printError('ERROR: Unable to proceed pricing calculation for WO# ' + workOrderNumber + ' - PMA not allowed by customer\n')
        displayDivision()

        return 
    
    else:
        pricingMethod = wo._customer['PricingMethod']
        print('Customer\'s pricing method is \t' + str(pricingMethod))

    if pricingMethod == 'Contract':
        ContractPricing(wo).calculatePricing()

    elif pricingMethod == 'Agreement':
        AgreementPricing(wo).calculatePricing()
    
    elif pricingMethod == 'Standard':
        StandardPricing(wo).calculatePricing()

    else:
        StandardPricing(wo).calculatePricing()

    displayDivision()


    return 

def main():
    clear_screen()


    printColor('\n\t\tWelcome to HRD WO Pricing!', Fore.WHITE + Back.LIGHTCYAN_EX + Style.DIM)
    printColor('\n\t\t\t ~ a contribution to the future of CLM ~', Fore.WHITE + Back.LIGHTCYAN_EX + Style.DIM)


    startLoop = testConnections()


    while startLoop:
        resetColor()
        startPricing()
        
    
    restartConn = inputColor('\tRestart connection? (Y/N):\t').strip().upper()
    if restartConn == 'Y' or restartConn == 'y':
        main()

main()