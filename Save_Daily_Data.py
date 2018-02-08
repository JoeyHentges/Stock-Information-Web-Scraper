#The role of this script is to import all of the data into the Single Day data excel sheet
#It takes in data from the Excel_Daily_Updater.py script
import xlrd
from xlutils.copy import copy
import datetime

def start(row, stock, openPrice, averagePrice, endPrice, lowestPrice, lowestPriceTime, lowestPriceVolume, highestPrice, highestPriceTime, highestPriceVolume, week52Low, week52High, totalVolume, averageVolume, outstandingShares, marketCap, institutionalOwnership):
    excelSheet = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Single_Day_Data.xls' #File name
    book = xlrd.open_workbook(excelSheet) #Open the excel file
    sh = copy(book) #Copy the excel file to a writable doc
    sheet = sh.get_sheet(0) #Use sheet 1
    
    #Write the data to the excel sheet
    sheet.write(row,0,datetime.datetime.now().strftime("%m/%d/%Y")) #date

    sheet.write(row,1,openPrice)
    sheet.write(row,2,averagePrice)
    sheet.write(row,3,endPrice)
    sheet.write(row,4,lowestPrice)
    sheet.write(row,5,lowestPriceTime)
    sheet.write(row,6,lowestPriceVolume)
    sheet.write(row,7,highestPrice)
    sheet.write(row,8,highestPriceTime)
    sheet.write(row,9,highestPriceVolume)
    sheet.write(row,10,week52Low)
    sheet.write(row,11,week52High)
    sheet.write(row,12,totalVolume)
    sheet.write(row,13,averageVolume)
    sheet.write(row,14,outstandingShares)
    sheet.write(row,15,marketCap)
    sheet.write(row,16,institutionalOwnership)

    #Format the columns to 20 characters wide
    for i in range(17):
        sheet.col(i).width = 20*256
    sh.save(excelSheet) #Save the excel file

    print('Saved: '+stock.upper())
    print()
    
