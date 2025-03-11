from logData import *
from colorama import *
from displayColorama import *

class Validation:

    details = {'AcceptsPMAs':None, 'PMANotListed':[], 'PMASpecified':False } 
    wo ={}

    def __init__(self, wo):

        self.wo = wo
        self.wo._pricing_details['ValidationDetails'] =  self.details
    
    def getValidationDetails(self):
        ret = ""
        for key, value in self.details.items():
            ret = ret + f"{key}: {value}" + ' '
        
        return ret

    def checkForPMAs(self, passExpired=False):
        print('Checking for PMAs...\t' + self.wo._customer['AcceptsPMAs'])
        pmaOkay = False

        self.details['PMANotListed'] = []
        pma_not_listed = []

        self.details['AcceptsPMAs'] = self.wo._customer['AcceptsPMAs']

        if self.wo._customer['AcceptsPMAs'] != 'PMAs Allowed':
            if len(self.wo._bom_pma_list) != 0:

                if len(self.wo._customer['PMAList']) == 0: 
                    pma_not_listed = self.wo._bom_pma_list
                else:
                    for pma in self.wo._bom_pma_list:
                        if pma['PnmAutoKey'] not in self.wo._customer['PMAList']:
                            pma_not_listed.append(pma)
                
                self.details['PMANotListed'] = pma_not_listed
                if len(pma_not_listed) != 0:
                    if len(pma_not_listed) == 1: 
                        print('The following PMA was ', end='') 
                        printColor('not found', Fore.RED, endWith='')
                        print(' in the customer\'s approved list:')
                    else:
                        print('The following PMAs were ', end='') 
                        printColor('not found', Fore.RED, endWith='')
                        print(' in the customer\'s approved list:')

                    for p in pma_not_listed:
                        printColor('\tPN: ' + p['PartNumber'] + ' / ' + p['Description'] + '\t DISPOSITION: ' + p['Disposition'] + '\t' + p['BOMStatus'], Fore.RED)
                        
                    if passExpired:
                        pmaOkay = True

                    else:
                        pmaInPO = 'N'
                        
                        if len(pma_not_listed) == 1: 
                            pmaInPO = inputColor('Is the above PMA specified in the purchase order? (Y/N): ').strip().upper()
                        else:
                            pmaInPO = inputColor('Are all the above PMAs specified in the purchase order? (Y/N): ').strip().upper()

                        if pmaInPO == 'Y' or pmaInPO == 'y':
                            self.details['PMASpecified'] = True
                            pmaOkay = True

                    for pma in pma_not_listed:
                        logPMANotListed(self.wo._customer['CmpAutoKey'], pma['PnmAutoKey'], self.wo._wo['WooAutoKey'], pmaOkay)

                else: 
                    print('All WO\'s BOM PMAs allowed by the customer!')
                    pmaOkay = True
            else:
                print('No PMAs found in WO\'s BOM.')
                pmaOkay = True   
        else:
            print('Customer has no PMA restrictions.')
            pmaOkay = True


        self.wo._pricing_details['ValidationDetails'] =  self.details
        
        return pmaOkay

