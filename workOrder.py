from datetime import date, datetime, timedelta
from logData import *
from quantumConnection import *
from salesforceConnection import *
from displayColorama import *

class WorkOrder:
    _workOrderNumber = ''
    _wo = {'WorkOrder':'', 'WooAutoKey':0, 'PORONumber':'', 'WwtAutoKey':None, 'WorkType':None, 'PnmAutoKey':0, 'PartNumber':'', 'Description':'', 'PartsFlatPrice':None, 'LaborFlatPrice':None, 'MiscFlatPrice':None}
    _customer = {'CmpAutoKey':0, 'CompanyCode':'', 'CompanyName':'', 'AcceptsPMAs':'Not Specified', 'PMAList':[], 'PricingMethod':'Standard','CLMDetails':{}}
    _bom_list = []
    _bom_pma_list = []
    _pricing_details = {'ValidationDetails':{}, 'PricingDetails':{}}

    def __init__(self, work_order_number):
        
        self._workOrderNumber = work_order_number

        print('Gathering data...')

        setWO(self)

        if self._workOrderNumber != '': 
            printColor('WO#: ' + self._wo['WorkOrder'] + '\tPO/RO#: ' + self._wo['PORONumber'] + '\tCUSTOMER: ' + self._customer['CompanyCode'] + ' - ' + self._customer['CompanyName'] + '\tPART NUMBER: ' + self._wo['PartNumber'] + ' / ' + self._wo['Description'] + '\tWORK SCOPE: ' + str(self._wo['WorkType']), Back.LIGHTBLUE_EX + Fore.WHITE + Style.DIM)

            checkForBOMAWP(self)

def reset(self):
    self._workOrderNumber = ''
    self._wo = {}
    self._customer = {}
    self._bom_list = []
    self._bom_pma_list = []
    self = None

    return self

def setWO(self):
    SQL = "SELECT * FROM WO_OPERATION WHERE SI_NUMBER ='" + str(self._workOrderNumber) + "'"
    w = quantumConnectionFetchOne(SQL)
    
    if w is None:
        printError('ERROR: Unable to proceed - WO# \'' + str(self._workOrderNumber) + '\' not found\n')
        reset(self)
        return 
    
    self._wo['WorkOrder'] = str(w.SI_NUMBER)
    self._wo['WooAutoKey'] = int(w.WOO_AUTO_KEY)
    self._wo['PORONumber'] = str(w.COMPANY_REF_NUMBER)

    self._wo['WwtAutoKey'] = int(w.WWT_AUTO_KEY) if w.WWT_AUTO_KEY is not None else None
    self._wo['WorkType'] = quantumConnectionFetchOne("SELECT * FROM QCTL.WO_WORK_TYPE WHERE WWT_AUTO_KEY = " + str(self._wo['WwtAutoKey'] )).WORK_TYPE if w.WWT_AUTO_KEY is not None else None

    self._wo['PartsFlatPrice'] = round(float(w.PARTS_FLAT_PRICE),2)
    self._wo['LaborFlatPrice'] = round(float(w.LABOR_FLAT_PRICE),2)
    self._wo['MiscFlatPrice'] = round(float(w.MISC_FLAT_PRICE),2)

    self._wo['PnmAutoKey'] = int(w.PNM_AUTO_KEY)
    pn = getPN(self._wo['PnmAutoKey'])
    self._wo['PartNumber'] = pn.PN
    self._wo['Description'] = pn.DESCRIPTION


    self._customer['CmpAutoKey'] = int(w.CMP_AUTO_KEY)
    customer = getCustomer(self._customer['CmpAutoKey'])
    self._customer['CompanyCode'] = customer['COMPANY_CODE']
    self._customer['CompanyName'] = customer['COMPANY_NAME']
    self._customer['AcceptsPMAs'] = customer['AcceptsPMAs']
    self._customer['PMAList'] = customer['PMAList']
    self._customer['PricingMethod'] = customer['PricingMethod']
    self._customer['CLMDetails'] = customer['CLMDetails']

    self._bom_list = getBOM(self._wo['WooAutoKey'])

    self._bom_pma_list = getPMAList(self._bom_list)

    return
    
def getPN(pnm_auto_key):
    SQL = "SELECT * FROM QCTL.PARTS_MASTER WHERE PNM_AUTO_KEY = " + str(pnm_auto_key) 

    return quantumConnectionFetchOne(SQL)

def getCustomer(cmp_auto_key):
    customer = {'COMPANY_CODE':'', 'COMPANY_NAME':'', 'AcceptsPMAs':'Not Specified', 'PMAList':[], 'PricingMethod':'Standard', 'CLMDetails':{}}

    SQL = "SELECT * FROM QCTL.COMPANIES WHERE CMP_AUTO_KEY = " + str(cmp_auto_key) 
    fetchedCustomer = quantumConnectionFetchOne(SQL)

    customer['COMPANY_CODE'] = fetchedCustomer.COMPANY_CODE
    customer['COMPANY_NAME'] = fetchedCustomer.COMPANY_NAME

    SOQL = 'SELECT Accepts_PMAs__c FROM Account WHERE CMP_AUTO_KEY__c = ' + str(cmp_auto_key)
    fetchedCustomer = salesforceQueryAll(SOQL)
    if len(fetchedCustomer) != 0 :
        customer['AcceptsPMAs'] = fetchedCustomer['AcceptsPMAs'] if fetchedCustomer['AcceptsPMAs'] is not None else 'Not Specified'

    SOQL ='SELECT PMA_Part__r.PNM_AUTO_KEY__c FROM hrd_customer_pma__c WHERE Customer__r.CMP_AUTO_KEY__c = ' + str(cmp_auto_key)
    fetchedCustomer = salesforceQueryAll(SOQL)
    if len(fetchedCustomer) != 0 :
         customer['PMAList'] = fetchedCustomer['PMAPart.PNMAUTOKEY']

    SOQL = 'SELECT Pricing_Method__c FROM hrd_clm_contract_header__c WHERE Inactive__c = false AND Start_Date__c < TODAY AND Customer__r.CMP_AUTO_KEY__c = ' + str(cmp_auto_key)
    fetchedCustomer = salesforceQueryAll(SOQL)
    if len(fetchedCustomer) != 0 :
        customer['PricingMethod'] = fetchedCustomer['PricingMethod'] if fetchedCustomer['PricingMethod'] is not None else 'Standard'

    SOQL = 'SELECT Auto_Approve_Threshold__c, BER_Threshold__c, Labor_Rate__c, Mark_Up_Cap__c, Mark_Up_Percent__c  FROM hrd_customer_clm_details__c WHERE Customer__r.CMP_AUTO_KEY__c = ' + str(cmp_auto_key)
    fetchedCustomer = salesforceQueryAll(SOQL)
    if len(fetchedCustomer) != 0 :
        customer['CLMDetails'] = fetchedCustomer if fetchedCustomer is not None else {}

    return customer

def getBOM(woo_auto_key):
    SQL = "SELECT * FROM QCTL.WO_BOM WHERE WOO_AUTO_KEY = " + str(woo_auto_key) 

    fetchedBOM = quantumConnectionFetchAll(SQL)
    BOM_LIST = []

    for b in fetchedBOM:
        BOM = {'PnmAutoKey':0, 'PartNumber':'', 'Description':'', 'Disposition':'','BOMStatus':None, 'PMAFlag':False, 'Qty':1, 'ListPrice':0.00, 'ListPriceDate':None, 'ListPriceDateAge':'Expired', 'KitComponents':[], 'OverwrittenListPrice': False}

        BOM['PnmAutoKey'] = int(b.PNM_AUTO_KEY)
        BOM['Qty'] = int(b.QTY_NEEDED)        
        BOM['BOMStatus'] = '*AWP*' if b.BOS_AUTO_KEY == 4 else ''

        PN = getPN(BOM['PnmAutoKey'])
        BOM['PartNumber'] = PN.PN
        BOM['Description'] = PN.DESCRIPTION
        BOM['PMAFlag'] = True if PN.PMA_FLAG == 'T' else False
        BOM['ListPrice'] = round(float(PN.LIST_PRICE),2)
        BOM['ListPriceDate'] = (PN.LIST_PRICE_DATE).strftime("%m/%d/%Y") if PN.LIST_PRICE_DATE is not None else None
        BOM['ListPriceDateAge'] = dateAge(BOM['ListPriceDate'], BOM['PnmAutoKey'])

        DISPOSITION = ''
        DISPOSITION = quantumConnectionFetchOne("SELECT OVERRIDE FROM QCTL.VIEW_BOM_ACTIVITIES WHERE VALUE = '" + str(b.ACTIVITY) + "'").OVERRIDE
        
        if DISPOSITION == None:
            DISPOSITION = b.ACTIVITY

        BOM['Disposition'] = DISPOSITION

        if ('kit' in BOM['PartNumber'].lower()) or ('kit' in BOM['Description'].lower()):
            fetchedKit = quantumConnectionFetchAll("SELECT * FROM QCTL.KIT_COMPONENTS WHERE PNM_AUTO_KEY = " + str(BOM['PnmAutoKey']))

            for k in fetchedKit:
                kit = {'PnmAutoKey':0, 'PartNumber':'', 'Description':'', 'PMAFlag':False, 'Qty':1,'ListPrice':0.00, 'ListPriceDate':None, 'ListPriceDate':'Expired'}
                
                kit['PnmAutoKey'] = int(k.KIT_PNM_AUTO_KEY)
                kit['Qty'] = int(k.QTY_ITEM)

                kit_comp = getPN(k.KIT_PNM_AUTO_KEY)

                kit['PartNumber'] = kit_comp.PN
                kit['Description'] = kit_comp.DESCRIPTION
                kit['PMAFlag'] = True if kit_comp.PMA_FLAG == 'T' else False
                kit['ListPrice'] = round(float(kit_comp.LIST_PRICE),2)
                kit['ListPriceDate'] = (kit_comp.LIST_PRICE_DATE).strftime("%m/%d/%Y") if kit_comp.LIST_PRICE_DATE is not None else None
                kit['ListPriceDateAge'] = dateAge(kit['ListPriceDate'], kit['PnmAutoKey'])

                if kit_comp.PMA_FLAG == 'T': BOM['PMAFlag'] = True

                BOM['KitComponents'].append(kit)

            if BOM['ListPriceDateAge'] == 'Expired':
                notExpiredKitListPrice = True
                for k in BOM['KitComponents']:
                    if ('ListPriceDateAge', 'Expired') in k.items():
                        notExpiredKitListPrice = False
                        
                    if not notExpiredKitListPrice: break
                
                if notExpiredKitListPrice:
                    minListPriceDate = None
                    sumListPriceDate = 0

                    if len(BOM['KitComponents']) != 0:
                        minListPriceDate =min([k['ListPriceDate'] for k in BOM['KitComponents']])
                        sumListPriceDate = sum([(k['Qty']*k['ListPrice'] ) for k in BOM['KitComponents']])
                    
                    BOM['OverwrittenListPrice'] = True
                    BOM['ListPrice'] = round(float(sumListPriceDate),2)
                    BOM['ListPriceDate'] = minListPriceDate 
                    BOM['ListPriceDateAge'] = dateAge(minListPriceDate, BOM['PnmAutoKey'])

        BOM_LIST.append(BOM)

    return BOM_LIST

def getPMAList(bom_list):
    PMA_LIST = []
    for bom in bom_list:
        pma = [bom for k, v in bom.items() if k =='PMAFlag' and v == True ]
        if len(pma) != 0: 
            for p in pma:
                PMA_LIST.append(p)

    return PMA_LIST

def dateAge(dateValue, PNM_AUTO_KEY):
    age = 'Expired'
    if dateValue is not None:
        dateValue = datetime.strptime(dateValue,"%m/%d/%Y")
        if isinstance(dateValue, (date, datetime)):
            todayYear = (date.today()).year
            dateYear = dateValue.strftime("%Y")
            yearsAgo = (int(todayYear) - int(dateYear))

            if yearsAgo in range(0, 2):
                age = 'Good'

            elif yearsAgo in range(2, 4):
                age = 'Acceptable' 

            else:
                age = 'Expired'

        if age == 'Expired' and PNM_AUTO_KEY > 0 :
            logPNExpiredListPriceDate(dateValue, PNM_AUTO_KEY)

    return age

def checkForBOMAWP(self):
    first = True
    nextDeliveryDate = []
    for bom in self._bom_list:
        for key, value in bom.items():
            if key == 'BOMStatus' and value == '*AWP*':
                NEXT_DELIVERY_DATE = None
                SQL = "SELECT QCTL.PO_DETAIL.NEXT_DELIVERY_DATE FROM QCTL.PO_HEADER JOIN QCTL.PO_DETAIL ON QCTL.PO_HEADER.POH_AUTO_KEY = QCTL.PO_DETAIL.POH_AUTO_KEY WHERE QCTL.PO_DETAIL.NEXT_DELIVERY_DATE IS NOT NULL AND QCTL.PO_HEADER.OPEN_FLAG = 'T' AND QCTL.PO_DETAIL.PNM_AUTO_KEY = " + str(bom['PnmAutoKey']) 
                po = quantumConnectionFetchOne(SQL)
                if po is not None: NEXT_DELIVERY_DATE = (po.NEXT_DELIVERY_DATE).strftime("%m/%d/%Y") if po.NEXT_DELIVERY_DATE is not None else None

                if first:
                    printColor('\n\t\t***WARNING: THE FOLLOWING BOM PNs ARE \'AWP\':', Back.CYAN + Fore.WHITE)
                    first = False

                printColor('\t\t\tPN: ' + bom['PartNumber'] + ' / ' + bom['Description'] + '\t DISPOSITION: ' + bom['Disposition'] + '\tNEXT DELIVERY DATE: ' + str(NEXT_DELIVERY_DATE), Back.CYAN + Fore.WHITE)
                nextDeliveryDate.append(NEXT_DELIVERY_DATE)
    
    if (not first) and nextDeliveryDate:
        newLeadingTime = '+ 5 DAYS'
        if max(nextDeliveryDate) is not None:
            newLeadingTime = (datetime.strptime(max(nextDeliveryDate), "%m/%d/%Y")  + timedelta(days=5)).strftime("%m/%d/%Y")
        printColor('\t\tPLEASE MODIFY ORDER\'S LEADING TIME TO ', Back.CYAN + Fore.WHITE, endWith='')
        printColor('\t' + str(newLeadingTime) +'\n' , Back.CYAN + Fore.RED)