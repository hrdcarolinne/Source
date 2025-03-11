from logData import *
from quantumConnection import *
from salesforceConnection import *
from displayColorama import *
from colorama import *

class AgreementPricing:
    details = {'WorkOrder':'', 'ExpiredBOMPN':[],
           'LaborRate': 0, 'LaborHours':0, 'LaborFlatPrice':0, 'QuantumLaborFlatPrice':0, 'LaborPriceDifference':0, 'LaborPriceDifferencePercentage':0,
           'PartsFlatPrice':0, 'QuantumPartsFlatPrice':0, 'PartsPriceDifference':0, 'PartsPriceDifferencePercentage':0 }

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
        print('Calculating Agreement Pricing...')
        self.details['WorkOrder'] = self.wo._wo['WorkOrder']

        printError('ERROR: Unable to proceed pricing calculation for WO# ' + self.wo._workOrderNumber + ' - Agreement Pricing not available\n')   

        self.wo._pricing_details['PricingDetails'] =  self.details
