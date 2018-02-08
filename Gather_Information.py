#Script used to gather excel data from the Stock Full Data files
import xlrd
import Gather_Information_Functions


#Function to start this script
def start(stock, rowBeginning, endRow):
    #Excel file path for day information
    fileName = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Full_Data.xls'
    #Open the excel file to sheet 0
    book = xlrd.open_workbook(fileName)
    sheet = book.sheet_by_name('Sheet')

    #Start gathering information
    #Use this array for price related values
    currentPriceArray = sheet.col_values(2, rowBeginning-1, endRow-1)
    #rowBeginning-1
    #Row the lowest price occurs
    lowestPriceRow = 0
    #Row the highest price occurs
    highestPriceRow = 0

    #Use this array for time related values
    currentTimeArray = sheet.col_values(1, rowBeginning-1, endRow-1)

    #Use this array for volume related vlaues
    currentVolumeArray = sheet.col_values(8, rowBeginning-1, endRow-1)

    #Print basic information
    print('Stock: '+stock)
    print('Row Beggining: '+str(rowBeginning))
    print('Row End: '+str(endRow))

    #Get the stocks opening price
    openPrice = sheet.cell(rowBeginning-1,5).value
    print('Opening Price: '+str(openPrice))

    #Get the average price throughout the day
    averagePrice = Gather_Information_Functions.getAverage(currentPriceArray)
    print('Average price: '+str(averagePrice))

    #Get the last price a stock was during the day
    lastValue = len(currentPriceArray)
    endPrice = currentPriceArray[lastValue-1]
    print('The Last Price: '+str(endPrice))

    #Get the lowest price of the stock
    lowestPrice = Gather_Information_Functions.getLowestPrice(currentPriceArray)
    print('Lowest Price: '+str(lowestPrice))

    #Get the time when the lowest price first showed
    lowestPriceTime, lowestPriceRow = Gather_Information_Functions.getLowestPriceTime(lowestPrice, currentPriceArray, currentTimeArray)
    print('Lowest Price Time: '+str(lowestPriceTime))

    #Get the volume of stocks traded when the lowest price came up
    lowestPriceVolume = Gather_Information_Functions.getLowestPriceVolume(lowestPriceRow, currentVolumeArray)
    print('Lowest Price Volume: '+str(lowestPriceVolume))

    #Get the highest price of the stock
    highestPrice = Gather_Information_Functions.getHighestPrice(currentPriceArray)
    print('Highest Price: '+str(highestPrice))

    #Get the time when the highest price first showed
    highestPriceTime, highestPriceRow = Gather_Information_Functions.getHighestPriceTime(highestPrice, currentPriceArray, currentTimeArray)
    print('Highest Price Time: '+str(highestPriceTime))

    #Get the volume of stocks traded when the highest price came up
    highestPriceVolume = Gather_Information_Functions.getHighestPriceVolume(highestPriceRow, currentVolumeArray)
    print('Highest Price Volume: '+str(highestPriceVolume))

    #Get the 52 week low of the stock
    week52Low = sheet.cell_value(rowBeginning-1, 10)
    print('52 Week Low: '+str(week52Low))

    #Get the 52 week high of the stock
    week52High = sheet.cell_value(rowBeginning-1, 11)
    print('52 Week High: '+str(week52High))

    #Get the end total volume traded of the stock
    totalVolume = sheet.cell_value(endRow-1, 8)
    print('Total Volume Traded: '+str(totalVolume))

    #Get the average total volume for the stock traded
    averageVolume = sheet.cell_value(rowBeginning-1, 9)
    print('Average Volume Traded: '+str(averageVolume))

    #Get the outstanding shares of the stock
    outstandingShares = sheet.cell_value(rowBeginning-1, 16)
    print('Outstanding Shares: '+str(outstandingShares))

    #Get the market cap for the stock
    marketCap = sheet.cell_value(rowBeginning-1, 17)
    print('Market Cap: '+str(marketCap))

    #Get the institutional ownership of the stock
    institutionalOwnership = sheet.cell_value(rowBeginning-1, 19)
    print('Institutional Ownership: '+str(institutionalOwnership))
    
    

    print()
    print('-------------------------------------------------------------')
    print()

    return (openPrice, averagePrice, endPrice, lowestPrice, lowestPriceTime, lowestPriceVolume,highestPrice, highestPriceTime, highestPriceVolume, week52Low, week52High, totalVolume, averageVolume, outstandingShares, marketCap, institutionalOwnership)
    
