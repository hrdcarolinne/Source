import os
import pyodbc
import traceback
from datetime import date, datetime
from displayColorama import *
from colorama import *

def testAccessConnection():
    getAccessDriver()
    try: 
        print('\tMS Access Backend', end='')

        accessConnection = getAccessConnection()

        printColor(' \u2713', Fore.GREEN + Style.BRIGHT, endWith='\t')

        accessConnection.close()

        return True
    
    except Exception as e:
        error_string = "ERROR: MS Access Connection \n" + str(e) + ":\t" + str(traceback.format_exc()) + "\n\n" 
        printError('\t' + error_string)
        return False
    
def getAccessDriver():
    return [item for item in pyodbc.drivers() if item.startswith('Microsoft Access Driver')][0]


def getAccessConnection():
    connectionString = 'Driver={' + getAccessDriver()  + '};' r'DBQ=\\radon\Stuff\Tools\HRDWOPricing\Source\Backend\HRDWOPricing_be.mdb;'

    try:
        accessConnection  = pyodbc.connect(connectionString)

    except Exception as e:
        errorString = "ERROR: " + str(e) + "\n" + str(traceback.format_exc())
        printError(errorString)
    
    finally:
        return accessConnection
    
def logPNExpiredListPriceDate(dateValue, PNM_AUTO_KEY):
    PriceListDate = '#' + dateValue.strftime("%m/%d/%Y") + '#' if isinstance(dateValue, (date, datetime)) else 'NULL'

    insertAccessSQL = 'INSERT INTO tblPNExpiredListPriceDate(PNM_AUTO_KEY, ListPriceDate, Username) VALUES(' + str(PNM_AUTO_KEY) + ', ' + str(PriceListDate) + ', \''  + os.getlogin() +'\')'

    accessConnection = getAccessConnection()
    accessCursor = accessConnection.cursor()

    try:
        accessCursor.execute(insertAccessSQL)
        accessCursor.commit()
    
    except Exception as e:
        errorString = "ERROR: " + str(e) + "\n" + str(traceback.format_exc())
        print(errorString)
    
    finally:
        accessConnection.close()

def logPMANotListed(CMP_AUTO_KEY, PNM_AUTO_KEY, WOO_AUTO_KEY,POSpecified):
    insertAccessSQL = 'INSERT INTO tblPMANotListed(CMP_AUTO_KEY,PNM_AUTO_KEY, WOO_AUTO_KEY, POSpecified, Username) VALUES('+ str(CMP_AUTO_KEY) + ',' + str(PNM_AUTO_KEY) + ', ' + str(WOO_AUTO_KEY) + ', ' + str(POSpecified)  + ', \''  + os.getlogin() +'\')'

    accessConnection  = getAccessConnection()
    accessCursor = accessConnection.cursor()

    try:
        accessCursor.execute(insertAccessSQL)
        accessCursor.commit()
    
    except Exception as e:
        errorString = "ERROR: " + str(e) + "\n" + str(traceback.format_exc())
        print(errorString)
    
    finally:
        accessConnection.close()

