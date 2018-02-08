#Script used to create and manage the excel file for daily data
import os
import xlwt
import xlrd
import Gather_Information
import Save_Daily_Data


#If an excel file is not available, create one with titles in row 1
def createExcelFile(stock):
    #Excel file path
    excelSheet = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Single_Day_Data.xls'
    book = xlwt.Workbook()

    #If Not check for path existance
    if not os.path.exists(excelSheet):
        #os.makedirs(os.path.dirname(excelSheet))
        sheet = book.add_sheet('Sheet')
        sheet.write(0,0,'Date')
        sheet.write(0,1,'Opening Price')
        sheet.write(0,2,'Average Price')
        sheet.write(0,3,'Ending Price')
        sheet.write(0,4,'Lowest Price')
        sheet.write(0,5,'Lowest Price Time')
        sheet.write(0,6,'Lowest Price Current Volume Traded')
        sheet.write(0,7,'Highest Price')
        sheet.write(0,8,'Highest Price Time')
        sheet.write(0,9,'Highest Price Current Volume Traded')
        sheet.write(0,10,'52 Week Low Price')
        sheet.write(0,11,'52 Week High Price')
        sheet.write(0,12,'Total Volume Traded')
        sheet.write(0,13,'Average Volume Traded')
        sheet.write(0,14,'Outstanding Shares')
        sheet.write(0,15,'Market Cap')
        sheet.write(0,16,'Institutional Ownership')
        book.save(excelSheet)

    #Get the number of rows currently used in the excel sheet
    #Used for placing new data
    workbookOpen = xlrd.open_workbook(excelSheet)
    sheets = workbookOpen.sheets()
    rows = sheets[0].nrows #Number of rows in the excel sheet
    return rows

#Get the ending row for a stock
def getEndRow(stock):
    #Excel file path for day information
    excelSheet = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Full_Data.xls'
    
    #Get the number of rows currently used in the excel sheet
    #Used for placing new data
    workbookOpen = xlrd.open_workbook(excelSheet)
    sheets = workbookOpen.sheets()
    rows = sheets[0].nrows #Number of rows in the excel sheet

    return rows

#Function to start this script
def start(stock, rowBeginning):

    #Call the 'createExcelFile' function
    row = createExcelFile(stock)

    #Call the 'getEndRow' function
    endRow = getEndRow(stock)

    #Call the 'gatherInformation' function, and have it return all data
    openPrice = 0.0
    averagePrice = 0.0
    endPrice = 0.0
    lowestPrice = 0.0
    lowestPriceTime = 0.0
    lowestPriceVolume = 0.0
    highestPrice = 0.0
    highestPriceTime = 0.0
    highestPriceVolume = 0.0
    week52Low = 0.0
    week52High = 0.0
    totalVolume = 0.0
    averageVolume = 0.0
    outstandingShares = 0.0
    marketCap = 0.0
    institutionalOwnership = 0.0

    Gather_Information.start(stock, rowBeginning, endRow)

    openPrice, averagePrice, endPrice, lowestPrice, lowestPriceTime, lowestPriceVolume, highestPrice, highestPriceTime, highestPriceVolume, week52Low, week52High, totalVolume, averageVolume, outstandingShares, marketCap, institutionalOwnership = Gather_Information.start(stock, rowBeginning, endRow)

    Save_Daily_Data.start(row, stock, openPrice, averagePrice, endPrice, lowestPrice, lowestPriceTime, lowestPriceVolume, highestPrice, highestPriceTime, highestPriceVolume, week52Low, week52High, totalVolume, averageVolume, outstandingShares, marketCap, institutionalOwnership)



