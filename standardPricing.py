from logData import *
from quantumConnection import *
from salesforceConnection import *
from displayColorama import *
from colorama import *

defaultLaborRate = 112
defaultMarkUpPercentage = 0.2
defaultAgeMarkUpPercentage = 0.1
dispositionToCharge = ('MISSING', 'Replace', 'CID', 'Consumable', 'EXCLUSION', 'Repair', 'OSP')

class StandardPricing:

    details = {'WorkOrder':'', 'ExpiredBOMPN':[],
           'LaborRate': defaultLaborRate, 'LaborHours':0, 'LaborFlatPrice':0, 'QuantumLaborFlatPrice':0,
           'PartsFlatPrice':0, 'QuantumPartsFlatPrice':0, 'QuantumMiscFlatPrice':0, 'QuantumPriceTotal':0, 'PriceTotal':0, 'PriceDifference':0, 'PriceDifferencePercentage':0 }

    wo ={}

    def __init__(self, wo):

        self.wo = wo
        self.wo._pricing_details['PricingDetails'] =  self.details

    def getPricingDetails(self):
        ret = ""
        for key, value in self.details.items():
            ret = ret + f"{key}: {value}" + ' '
        
        return ret

    def calculatePricing(self, passValidation=False):

        print('Calculating Standard Pricing...')
        self.details['WorkOrder'] = self.wo._wo['WorkOrder']

        if self.checkUsableBOMPrice() or passValidation:
            printColor('CALCULATED PRICES:', Back.LIGHTMAGENTA_EX + Fore.WHITE)
            LaborFlatPrice = self.computeLaborFlatRate()
            PartsFlatPrice = self.computePartsFlatRate()

            priceTotal = self.details['LaborFlatPrice'] + self.details['PartsFlatPrice']
            self.details['PriceTotal'] = round(priceTotal,2)

            self.details['QuantumMiscFlatPrice'] = self.wo._wo['MiscFlatPrice']

            quantumTotal = round(self.details['QuantumLaborFlatPrice'] + self.details['QuantumPartsFlatPrice'] + self.details['QuantumMiscFlatPrice'],2)
            self.details['QuantumPriceTotal'] = round(quantumTotal,2)

            self.details['PriceDifference'] = round(priceTotal - quantumTotal,2)
            self.details['PriceDifferencePercentage'] = round(((priceTotal - quantumTotal)/priceTotal)*100,2) if priceTotal != 0 else 0
            
            
            printColor('\n\t\tPRICE TOTAL = \t$' + str(self.details['PriceTotal']), Fore.WHITE + Back.LIGHTMAGENTA_EX) 

            printColor('\nQUOTED IN Q FOR THE WORK ORDER: \t [Labor = $' + str(self.details['QuantumLaborFlatPrice']) + ',\t Parts = $'+ str(self.details['QuantumPartsFlatPrice']) + ',\t Misc = $' + str(self.details['QuantumMiscFlatPrice']) + ']\t=\t$' + str(self.details['QuantumPriceTotal']) ,Fore.MAGENTA)
            
            printColor('\n\t\tPRICE DIFFERENCE  = ',Fore.MAGENTA, endWith='')
            printColor('$' + str(self.details['PriceDifference']), Fore.WHITE + Back.LIGHTMAGENTA_EX, endWith='') 
            printColor('\t-->\t' + str(self.details['PriceDifferencePercentage']) + '%', Fore.MAGENTA)


        else:
            printError('ERROR: Unable to proceed pricing calculation for WO# ' + self.wo._workOrderNumber + ' - Expired list pricing\n')   

        self.wo._pricing_details['PricingDetails'] =  self.details

    def checkUsableBOMPrice(self):
        EXPIRED_LIST = []
        for bom in self.wo._bom_list:
            if bom['Disposition'] in dispositionToCharge:
                exp = [bom for k, v in bom.items() if k =='ListPriceDateAge' and v == 'Expired']
                if len(exp) != 0: 
                    for e in exp:
                        EXPIRED_LIST.append(e)
        
        self.details['ExpiredBOMPN'] = EXPIRED_LIST

        if len(EXPIRED_LIST) != 0: 
            if len(EXPIRED_LIST) == 1: 
                print('The PN below has ', end='') 
                printColor('Expired', Fore.RED, endWith='')
                print(' list price:')
            else:
                print('The PNs below has ', end='') 
                printColor('Expired', Fore.RED, endWith='')
                print(' list prices:')

            for b in EXPIRED_LIST:
                print('\tPN: ' + b['PartNumber'] + ' (' + b['Description'] + ') \t DISPOSITION: ' + b['Disposition'], end='')
                printColor('\t' + str(b['BOMStatus']), Fore.RED)
                print('\t\tLIST PRICE: $' + str(b['ListPrice']), end='') 
                printColor('\tDATE: ' + str(b['ListPriceDate']), Fore.RED)

            self.wo._pricing_details['PricingDetails'] =  self.details
            return False
        
        else:
            self.wo._pricing_details['PricingDetails'] =  self.details
            return True

    def computeLaborFlatRate(self):
        rate = defaultLaborRate
        if self.wo._customer['CLMDetails'] != {}: rate = self.wo._customer['CLMDetails']['LaborRate'] if 'LaborRate' in self.wo._customer['CLMDetails'] else defaultLaborRate

        laborHours = self.getLaborHours()
        
        if laborHours is None: return None

        total = round(laborHours*rate,2)

        printColor('\n\tLABOR TOTAL = ',Fore.MAGENTA, endWith='')
        printColor('$' + str(total), Fore.WHITE + Back.LIGHTMAGENTA_EX, endWith='') 
        # print('\t\t\t[labor hours: ' + str(laborHours) + ' X rate: ' + str(rate)+ ']')
        # print('\t\t\t[Quantum:  $' + str(self.wo._wo['LaborFlatPrice']))
        # print('\t\t\t\tprice difference:  $' + str(round(total - self.wo._wo['LaborFlatPrice'],2)) + ']')
        
        self.details['LaborRate'] = rate
        self.details['LaborHours'] = laborHours
        self.details['LaborFlatPrice'] = total
        self.details['QuantumLaborFlatPrice'] = self.wo._wo['LaborFlatPrice']
        self.wo._pricing_details['PricingDetails'] =  self.details

        return total


    def getLaborHours(self):
        workType = self.wo._wo['WorkType']
        if workType == 'OVERHAULED':
            fetchedLaborHours = quantumConnectionFetchOne('SELECT sum(qctl.wo_task_skills.est_hours) AS labor_est_hours FROM qctl.wo_operation INNER JOIN qctl.wo_task on qctl.wo_task.woo_auto_key = qctl.wo_operation.woo_auto_key INNER JOIN qctl.wo_task_skills on qctl.wo_task.wot_auto_key = qctl.wo_task_skills.wot_auto_key WHERE qctl.wo_task.wtm_auto_key <> 556 AND qctl.wo_operation.woo_auto_key = '+ str(self.wo._wo['WooAutoKey']) + ' GROUP BY qctl.wo_operation.woo_auto_key;')
            if fetchedLaborHours is not None:
                if len(fetchedLaborHours) != 0:
                    return fetchedLaborHours[0] if fetchedLaborHours[0] is not None else 0
                else:
                    return 0
            else:
                return 0
        else:
            printError('Labor hours for ' + workType + ' not available' )
                
            return None
        
    def computePartsFlatRate(self):
        sumTotal = 0
        print('\n\tLIST PRICE + (LIST PRICE X MARK UP %)  + (LIST PRICE X AGE MARK UP %) \t x QTY\n')

        for pn in self.wo._bom_list:
            if pn['Disposition'] in dispositionToCharge:
                markUpPercentage = defaultMarkUpPercentage
                ageMarkUpPercentage = defaultAgeMarkUpPercentage if pn['ListPriceDateAge'] == 'Acceptable' else 0
                listPrice = pn['ListPrice']
                qty = pn['Qty']

                print('PN: ' + pn['PartNumber'] + ' / ' + pn['Description'], end='')
                printColor(str('*' if pn['OverwrittenListPrice'] else ''), Fore.RED, endWith='')
                print('\tQTY:' + str(qty)+ '\t DISPOSITION: ' + pn['Disposition'] , end='')
                printColor('\t' + str(pn['BOMStatus']), Fore.RED)

                total = listPrice
                total += (listPrice*markUpPercentage) 
                total += (listPrice*ageMarkUpPercentage)
                total = round((total*qty),2)

                print('\t$' + str(listPrice) + ' + ($' + str(listPrice) + ' X ' +  str(markUpPercentage) + ') + ($' + str(listPrice) + ' X ' +  str(ageMarkUpPercentage) + ')\t x ' + str(qty) + '\t = $' + str(total))
                
                sumTotal += total

            sumTotal = round(sumTotal,2)

        printColor('\n\tPARTS TOTAL = ',Fore.MAGENTA, endWith='')
        printColor('$' + str(sumTotal), Fore.WHITE + Back.LIGHTMAGENTA_EX) 
        # print('\t\t\t[Quantum:  $' + str(self.wo._wo['PartsFlatPrice']))
        # print('\t\t\t\tprice difference:  $' + str(round(sumTotal - self.wo._wo['PartsFlatPrice'],2)) + ']')

        sumTotal = round(sumTotal,2)

        self.details['PartsFlatPrice'] = sumTotal
        self.details['QuantumPartsFlatPrice'] = self.wo._wo['PartsFlatPrice']

        self.wo._pricing_details['PricingDetails'] =  self.details

        return sumTotal