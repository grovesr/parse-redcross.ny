#!/usr/bin/python
"""
Created by: Rob Groves
Date: 07/24/15

parse NYC Red Cross inventory spreadsheets into a format suitable
for uploading to RIMS database 
"""
import sys
import os
import glob
import xlwt
from xlrdutils import xlrdutils
import code
from optparse import OptionParser

def get_sites(filename =None):
    if not filename:
        return {}
    try:
        workbook=xlrdutils.open_workbook(filename=filename)
    except (xlrdutils.XlrdutilsOpenWorkbookError,
            xlrdutils.XlrdutilsOpenSheetError) as e:
        warningMessage = repr(e)
        print warningMessage
        sys.exit(-1)
    try:
        data=xlrdutils.read_lines(workbook, 
                                  headerKeys=['Site Number',
                                              'Site Name',
                                              'RC Site Name',],)
    except (xlrdutils.XlrdutilsReadHeaderError,
            xlrdutils.XlrdutilsDateParseError) as e:
        warningMessage = repr(e)
        print warningMessage
        raise e
        sys.exit(-1)
    siteDict={}
    rcSiteNameList = data['RC Site Name']
    siteNameList = data['Site Name']
    siteNumberList = data['Site Number']
    
    for k in range(len(rcSiteNameList)):
        if 'Totals' not in rcSiteNameList[k] and rcSiteNameList[k] != '':
            try:
                siteDict[rcSiteNameList[k]]=(siteNameList[k],siteNumberList[k])
            except KeyError as e:
                print repr(e)
                raise e
                sys.exit(-1)
    return siteDict

def get_products(filename =None):
    print filename
    if not filename:
        return {}
    try:
        workbook=xlrdutils.open_workbook(filename=filename)
    except (xlrdutils.XlrdutilsOpenWorkbookError,
            xlrdutils.XlrdutilsOpenSheetError) as e:
        warningMessage = repr(e)
        print warningMessage
        raise e
        sys.exit(-1)
    try:
        data=xlrdutils.read_lines(workbook, 
                                  headerKeys=['Product Code',
                                              'Unit of Measure',
                                              'Qty of Measure',],)
    except (xlrdutils.XlrdutilsReadHeaderError,
            xlrdutils.XlrdutilsDateParseError) as e:
        warningMessage = repr(e)
        print warningMessage
        raise e
        sys.exit(-1)
    productDict={}
    productCodeList = data['Product Code']
    qtyOfMeasureList = data['Qty of Measure']
    
    for k in range(len(productCodeList)):
        try:
            productDict[productCodeList[k]]=qtyOfMeasureList[k]
        except KeyError as e:
            print repr(e)
            raise e
            sys.exit(-1)
    return productDict

def parse_sites(rcSiteNameList,deliverySiteFilename):
    siteDict = get_sites(filename=deliverySiteFilename)
    siteList=[]
    for siteName in rcSiteNameList:
        if 'Totals' not in siteName and siteName != '':
            siteList.append(siteDict[unicode(siteName.encode('ascii','replace').replace('?','-'))])
    return siteList

def create_inventory_workbook(data):
    xls = xlwt.Workbook(encoding="utf-8")
    sheet1 = xls.add_sheet("Inventory")
    sheet1.write(0,0,'Product Code')
    sheet1.write(0,1,'Prefix')
    sheet1.write(0,2,'Site Number')
    sheet1.write(0,3,'Cartons')
    rowIndex = 1
    for code,siteInventoryList in data.iteritems():
        for item in siteInventoryList:
            sheet1.write(rowIndex,0,code)
            sheet1.write(rowIndex,1,'P')
            sheet1.write(rowIndex,2,int(item[1]))
            sheet1.write(rowIndex,3,int(item[2]))
            rowIndex += 1
    return xls

def calculate_pkg_qty(data, productInformationFilename):
    """
    calculate package quantities from per piece quantities
    """
    # data keys are product codes
    # values are tuples (siteName, siteNumber, siteQuantity)
    productDict = get_products(filename=productInformationFilename)
    extendedData = {}
    for code, siteInventoryList in data.iteritems():
        try:
            divisor = productDict[code]
        except KeyError as e:
            print repr(e)
            print productDict
            raise e
            sys.exit(-1)
        for item in siteInventoryList:
            if code not in extendedData:
                extendedData[code] = []
            qty = item[2]
            if qty != '' and qty != 'na' and qty != 'x' and qty != 'X':
                extendedData[code].append((item[0],
                                           item[1],
                                           int(qty) / int(divisor),))
        if len(extendedData[code]) == 0:
            del extendedData[code]
    return extendedData

def main():
    # begin main program
    usage = "usage: %prog -d INVENTORYDIR"
    parser = OptionParser(usage)
    parser.add_option("-d", "--dir", dest="inventoryDirFullPathName",
                      help="read inventory spreadsheets in INVENTORYDIR")
    (options, args) = parser.parse_args()
    if not options.inventoryDirFullPathName:
        parser.error("No -d INVENTORYDIR supplied")
    # determine inventory directory name
    inventoryDirFullPathName = options.inventoryDirFullPathName
    dirContents = glob.glob(inventoryDirFullPathName + os.sep + '*inventory*.xls')
    deliverySiteFilename = inventoryDirFullPathName + os.sep + 'Delivery_Sites.xls'
    productInformationFilename = inventoryDirFullPathName + os.sep + 'Product_Information_Each.xls'
    dirContents.sort()
    allData = {}
    sheets = ['DS Supplies',
              'Food Related',
              'Clothing',
              'Other',]
    for filename in dirContents:
        print filename
        try:
            workbook=xlrdutils.open_workbook(filename=filename)
        except (xlrdutils.XlrdutilsOpenWorkbookError,
                xlrdutils.XlrdutilsOpenSheetError) as e:
            warningMessage = repr(e)
            print warningMessage
            raise e
            continue
        for k in range(len(sheets)):
            sheetName = sheets[k]
            try:
                data=xlrdutils.read_lines(workbook, 
                                          headerKeys=['Location.*',],
                                          sheet=sheetName,)
            except (xlrdutils.XlrdutilsReadHeaderError,
                    xlrdutils.XlrdutilsDateParseError) as e:
                warningMessage = repr(e)
                print warningMessage
                raise e
                sys.exit(-1)
            for header in data.keys():
                if 'Location' in header:
                    # returns list of tuples [(siteName, siteNumber),]
                        siteList = parse_sites(data[header],deliverySiteFilename)
            for headerVal,siteQtyList in data.iteritems():
                if 'Location' not in headerVal:
                    if headerVal not in allData:
                        allData[headerVal] = []
                    for k in range(len(siteList)):
                        # headerVal is the product code, so we are appending a
                        # three tuple (siteName, siteNumber, product quantity) to
                        # a list of three tuples for a given product code
                        allData[headerVal].append((siteList[k][0],siteList[k][1],siteQtyList[k]))
    # this puts out package counts
    extendedData = calculate_pkg_qty(allData, productInformationFilename)
    # this puts out piece counts
    #extendedData = allData
    inventoryData={}
    for code,siteInventoryList in extendedData.iteritems():
        for inventoryTuple in siteInventoryList:
            if inventoryTuple[2] != '' and inventoryTuple[2] != 'na' and inventoryTuple[2] != 'x' and inventoryTuple[2] != 'X':
                if code not in inventoryData:
                    inventoryData[code] = []
                inventoryData[code].append(inventoryTuple)
    xls = create_inventory_workbook(inventoryData)
    xls.save(inventoryDirFullPathName + os.sep + 'Inventory.xls')

if __name__ == '__main__':
    main()