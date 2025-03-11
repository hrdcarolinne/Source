import os
import pandas as pd
from datetime import date, timedelta
from logData import *
from quantumConnection import *
from salesforceConnection import *
from workOrder import *
from validation import *
from standardPricing import *
from contractPricing import *
from agreementPricing import *
from displayColorama import *


dateFilter = date.today() - timedelta(days = 1)

_workOrder = []
_poRoNumber = []
_customerCode = []
_customerName = []
_workType = []
_partNumber = []
_acceptsPMAs = []
_pmaNotListed = []
_pmaSpecified  = []
_pricingMethod  = []
_expiredBOMPN = []
_laborRate = []
_laborHours = []
_laborFlatPrice = []
_quantumLaborFlatPrice = []
_partsFlatPrice = []
_quantumPartsFlatPrice = []
_quantumMiscFlatPrice = []
_quantumPriceTotal = []
_priceTotal = []
_priceDifference = []
_priceDifferencePercentage = []

def exportToExcel():
    clear_screen()
    SQL = "SELECT * FROM QCTL.AAA_CLM_HFCA_PRICECHECK WHERE QCTL.AAA_CLM_HFCA_PRICECHECK.ENTRY_DATE = \'" + dateFilter.strftime("%d-%b-%y") + "\'"

    printColor('Exporting HFCA WOs from ' + dateFilter.strftime("%d-%B-%y"), Back.LIGHTBLUE_EX + Fore.WHITE)

    fetchedWO = quantumConnectionFetchAll(SQL)

    for w in fetchedWO:
        wo = None

        printColor('Adding to list: WO# ' + w.SI_NUMBER, Fore.YELLOW + Back.LIGHTYELLOW_EX)

        wo = WorkOrder(w.SI_NUMBER)

        if wo._workOrderNumber != '': 
            
            getPricing(wo)

            addToSpreadsheet(wo)

        displayDivision()

    data = {'Work Order #':_workOrder, 'PO/RO #': _poRoNumber, 'Customer Code': _customerCode , "Customer Name": _customerName, 'Part Number': _partNumber, 'Work Type': _workType,
            'Accepts PMAs': _acceptsPMAs, 'PMA Not Listed': _pmaNotListed, 'PricingMethod': _pricingMethod, 'Expired BOM PN': _expiredBOMPN, 
            'Labor Rate': _laborRate, 'Labor Hours': _laborHours,'Calculated Labor Price': _laborFlatPrice,'Quantum Labor Flat Price': _quantumLaborFlatPrice,
            'Calculated Parts Price': _partsFlatPrice,'Quantum Parts Flat Price': _quantumPartsFlatPrice, 'Quantum Misc Flat Price': _quantumMiscFlatPrice,
            'Quantum Price Total': _quantumPriceTotal, 'Calculated Price Total': _priceTotal, 'Price Diff': _priceDifference, 'Price Diff %': _priceDifferencePercentage}
    
    df = pd.DataFrame(data)
    file_path = r'C:\Users\Carolinne\Desktop\HFCAPriceDiff_' + dateFilter.strftime("%Y-%m-%d") +'.xlsx'
    df.to_excel(file_path, index=False, sheet_name='OVERHAULED')

    printColor('Spreadsheet saved at ' + file_path, Fore.YELLOW + Back.LIGHTYELLOW_EX)

    os.system(f"start {file_path}")


def addToSpreadsheet(wo):
    workOrder = wo._wo['WorkOrder']
    pricingMethod = wo._customer['PricingMethod']
    validationDetails = wo._pricing_details['ValidationDetails']
    pricingDetails = wo._pricing_details['PricingDetails']

    pmaNotListed = ''
    expiredBOMPN = ''
    
    if pricingMethod =='Standard':
        if len(validationDetails['PMANotListed']) != 0:
            first = True
            for pma in validationDetails['PMANotListed']:
                if not first: pmaNotListed = pmaNotListed + ', \n'
                pmaNotListed = pmaNotListed + pma['PartNumber'] 
                first = False

        if len(pricingDetails['ExpiredBOMPN']) != 0:
            first = True
            for e in pricingDetails['ExpiredBOMPN']:
                if not first: expiredBOMPN = expiredBOMPN + ', \n'
                expiredBOMPN = expiredBOMPN + e['PartNumber'] + ' (' +  str(e['ListPriceDate']) + ')'
                first = False

        _workOrder.append(workOrder)
        _poRoNumber.append(wo._wo['PORONumber'])
        _customerCode.append(wo._customer['CompanyCode'])
        _customerName.append(wo._customer['CompanyName'])
        _partNumber.append(wo._wo['PartNumber'])
        _workType.append(wo._wo['WorkType'])
        _acceptsPMAs.append(validationDetails['AcceptsPMAs'])
        _pmaNotListed.append(pmaNotListed)
        _pmaSpecified.append(validationDetails['PMASpecified'])
        _pricingMethod.append(pricingMethod)
        _expiredBOMPN.append(expiredBOMPN)
        _laborRate.append(pricingDetails['LaborRate'])
        _laborHours.append(pricingDetails['LaborHours'])
        _laborFlatPrice.append(pricingDetails['LaborFlatPrice'])
        _quantumLaborFlatPrice.append(pricingDetails['QuantumLaborFlatPrice'])
        _partsFlatPrice.append(pricingDetails['PartsFlatPrice'])
        _quantumPartsFlatPrice.append(pricingDetails['QuantumPartsFlatPrice'])
        _quantumMiscFlatPrice.append(pricingDetails['QuantumMiscFlatPrice'])
        _quantumPriceTotal.append(pricingDetails['QuantumPriceTotal'])
        _priceTotal.append(pricingDetails['PriceTotal'])
        _priceDifference.append(pricingDetails['PriceDifference'])
        _priceDifferencePercentage.append(pricingDetails['PriceDifferencePercentage'])

def exportToAccessDBA():
    clear_screen()
    SQL = "SELECT * FROM QCTL.AAA_CLM_HFCA_PRICECHECK WHERE QCTL.AAA_CLM_HFCA_PRICECHECK.ENTRY_DATE = \'" + dateFilter.strftime("%d-%b-%y") + "\'"

    fetchedWO = quantumConnectionFetchAll(SQL)

    for w in fetchedWO:
        wo = None

        printColor(w.SI_NUMBER, Fore.YELLOW + Back.LIGHTYELLOW_EX)

        wo = WorkOrder(w.SI_NUMBER)

        if wo._workOrderNumber != '': 
            
            getPricing(wo)

            logPricing(wo)

        displayDivision()

def getPricing(wo):
    
    validationObj = Validation(wo)
    validationObj.checkForPMAs(True)

    pricingMethod = wo._customer['PricingMethod']

    pricingObj = None

    print('Customer\'s pricing method is \t' + str(pricingMethod))

    if pricingMethod == 'Contract':
        pricingObj = ContractPricing(wo)

    elif pricingMethod == 'Agreement':
        pricingObj = AgreementPricing(wo)
      
    elif pricingMethod == 'Standard':
        pricingObj = StandardPricing(wo)

    else:
        pricingObj = StandardPricing(wo)


    if pricingObj is not None:
        pricingObj.calculatePricing(True)

    return 
def logPricing(wo):
    workOrder = wo._wo['WorkOrder']
    pricingMethod = wo._customer['PricingMethod']
    validationDetails = wo._pricing_details['ValidationDetails']
    pricingDetails = wo._pricing_details['PricingDetails']

    pmaNotListed = ''
    expiredBOMPN = ''

    insertAccessSQL = None
    
    if pricingMethod =='Standard':
        if len(validationDetails['PMANotListed']) != 0:
            first = True
            for pma in validationDetails['PMANotListed']:
                if not first: pmaNotListed = pmaNotListed + ', '
                pmaNotListed = pmaNotListed + pma['PartNumber'] + '\n'
                first = False

        if len(pricingDetails['ExpiredBOMPN']) != 0:
            first = True
            for e in pricingDetails['ExpiredBOMPN']:
                if not first: expiredBOMPN = expiredBOMPN + ', '
                expiredBOMPN = expiredBOMPN + e['PartNumber'] + ' (' +  str(e['ListPriceDate']) + ')\n'
                first = False

        insertAccessSQL = "INSERT INTO tblStandardPricingLog ([WorkOrder], [PORONumber], [CustomerCode], [CustomerName], [WorkType], "
        insertAccessSQL = insertAccessSQL + "[AcceptsPMAs], [PMANotListed], [PMASpecified], [PricingMethod], [ExpiredBOMPN], "
        insertAccessSQL = insertAccessSQL + "[LaborRate], [LaborHours], [LaborFlatPrice], [QuantumLaborFlatPrice], "
        insertAccessSQL = insertAccessSQL + "[PartsFlatPrice], [QuantumPartsFlatPrice], [QuantumMiscFlatPrice], [PriceDifference], [PriceDifferencePercentage], [Username]) "
        insertAccessSQL = insertAccessSQL + "SELECT "
        insertAccessSQL = insertAccessSQL + "\'" + str(workOrder) + "\' AS [WorkOrder], "
        insertAccessSQL = insertAccessSQL + "\'" + str(wo._wo['PORONumber'])  + "\' AS [PORONumber], "
        insertAccessSQL = insertAccessSQL + "\'" + str(wo._customer['CompanyCode']) + "\' AS [CustomerCode], "
        insertAccessSQL = insertAccessSQL + "\'"  + str(wo._customer['CompanyName'])  + "\' AS [CustomerName], "
        insertAccessSQL = insertAccessSQL + "\'" + str(wo._wo['WorkType'])  + "\' AS [WorkType], "
        insertAccessSQL = insertAccessSQL + "\'"  + str(validationDetails['AcceptsPMAs']) + "\' AS [AcceptsPMAs], "
        insertAccessSQL = insertAccessSQL + "\'" + stringToSql(str(pmaNotListed)) + "\' AS [PMANotListed], "
        insertAccessSQL = insertAccessSQL + str(validationDetails['PMASpecified']) + " AS [PMASpecified], "
        insertAccessSQL = insertAccessSQL + "\'" + str(pricingMethod) + "\' AS [PricingMethod], "
        insertAccessSQL = insertAccessSQL + "\'"  + stringToSql(str(expiredBOMPN))  + "\' AS [ExpiredBOMPN], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['LaborRate']) + " AS [LaborRate], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['LaborHours']) + " AS [LaborHours], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['LaborFlatPrice']) + " AS [LaborFlatPrice], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['QuantumLaborFlatPrice']) + " AS [QuantumLaborFlatPrice], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['PartsFlatPrice']) + " AS [PartsFlatPrice], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['QuantumPartsFlatPrice']) + " AS [QuantumPartsFlatPrice], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['QuantumMiscFlatPrice']) + " AS [QuantumMiscFlatPrice], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['PriceDifference']) + " AS [PriceDifference], "
        insertAccessSQL = insertAccessSQL + str(pricingDetails['PriceDifferencePercentage']) + " AS [PriceDifferencePercentage], "
        insertAccessSQL = insertAccessSQL + "\'"  + os.getlogin() + "\' AS [Username]"

    accessConnection  = getAccessConnection()
    accessCursor = accessConnection.cursor()
    
    if insertAccessSQL is not None:
        try:
            accessCursor.execute(insertAccessSQL)
            accessCursor.commit()
        
        except Exception as e:
            errorString = "ERROR: " + str(e) + "\n" + str(traceback.format_exc())
            printError(errorString)
        finally:
            printColor('Pricing successfully logged',Back.LIGHTMAGENTA_EX + Fore.WHITE)

    accessConnection.close()

def stringToSql(stringToFormat):
    strFormated = str(stringToFormat).replace('\"', '""')
    strFormated = strFormated.replace("\'", "")

    ret = strFormated

    return ret

def clear_screen():
    os.system('cls||clear')
        
    resetColor()

exportToExcel()