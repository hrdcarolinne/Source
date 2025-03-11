from simple_salesforce import Salesforce
import traceback
from displayColorama import *
from colorama import *

sfUsername='carolinne.prod@hrd-aerosystems.com.hrdsb'
sfPassword='Hrd676985!'
sfSecurity_token='MRv2MTyBDFxPO5uGQkWUQXcKA'
sfDomain='test'


def testSalesforceConnection():
    try: 
        print('\tSalesforce Sandbox', end='')

        sfConnection = Salesforce(username=sfUsername, password=sfPassword, security_token=sfSecurity_token, domain=sfDomain)
        
        printColor(' \u2713', Fore.GREEN + Style.BRIGHT, endWith='\t')
        
        return True


    except Exception as e:
        error_string = "ERROR: Salesforce Connection \n" + str(e) + ":\t" + str(traceback.format_exc()) + "\n\n" 
        printError('\t' + error_string)

        return False

    
def salesforceQueryAll(SOQL):
    sf = Salesforce(username=sfUsername, password=sfPassword, security_token=sfSecurity_token, domain=sfDomain)
    records = {}

    results=sf.query_all_iter("" + SOQL + "")

    return getRecords(results)

def getRecords(results):
    records = {}
    for row in results:
        if isinstance(row, dict):
            row.pop('attributes', None)
            for key, value in row.items():
                if isinstance(value, dict):
                    value.pop('attributes', None)
                    for k, v in value.items():
                        if cropFieldNames(key + '.' + k) not in  records: 
                                records[cropFieldNames(key + '.' + k)] = []

                        records[cropFieldNames(key + '.' + k)].append(v)
                else:
                    records[cropFieldNames(key)] = value
    return records

def cropFieldNames(fieldName):
    fieldName = fieldName.replace('__c', '')
    fieldName = fieldName.replace('__r', '')
    fieldName = fieldName.replace('_', '')

    return fieldName

