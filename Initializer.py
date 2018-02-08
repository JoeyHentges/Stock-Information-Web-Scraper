import json
import requests
import os
import os.path
import xlwt
import xlrd
import Stock_Price_Grab    #Importing script <--
import Excel_Daily_Updater #Importing script <--
import INSTALLER #Import the installer for installing script packages
import datetime
import time

#Call the grabStock function from the Stock_Price_Grab script
def callStockGrab(stock):
    #Call the stock grab for 'stock'
    service = Stock_Price_Grab.grabStock(stock) #Returns whether the information was available or not
    return service


#Check for, and create the text file holding temperary information
def checkTextFilePath(stock):
    fileName = 'lastInformation_'+stock.upper()+'.txt'
    if not os.path.exists(fileName):
        file = open(fileName, 'w+')
        file.close()

#Remove the temperary text files
def removeTextFilePath(stock):
    fileName = 'lastInformation_'+stock.upper()+'.txt'
    if os.path.exists(fileName):
        os.remove(fileName)

#Check if starting excel file excists, create if it does not
def checkExcelFile(stock):
    #File location for the daily changing excel sheet
    excelSheet = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Full_Data.xls'
    book = xlwt.Workbook()
    
    #If the file does not exist, create the file with titles in row 0
    if not os.path.exists(excelSheet):
        os.makedirs(os.path.dirname(excelSheet))
        sheet = book.add_sheet('Sheet')
        sheet.write(0,0,'Date')
        sheet.write(0,1,'Time In Seconds')
        sheet.write(0,2,'Current Price')
        sheet.write(0,3,'Price Change')
        sheet.write(0,4,'Change Percentage')
        sheet.write(0,5,'Opening Price')
        sheet.write(0,6,'Lowest Price')
        sheet.write(0,7,'Highest Price')
        sheet.write(0,8,'Traded Today')
        sheet.write(0,9,'Daily Traded Average')
        sheet.write(0,10,'52 Week Low Price')
        sheet.write(0,11,'52 Week High Price')
        sheet.write(0,12,'Price/Earnings Ratio')
        sheet.write(0,13,'Earnings Per Share')
        sheet.write(0,14,'Yield')
        sheet.write(0,15,'Dividend')
        sheet.write(0,16,'Outstanding Shares')
        sheet.write(0,17,'Market Cap')
        sheet.write(0,18,'Beta')
        sheet.write(0,19,'Institutional Ownership')
        book.save(excelSheet)    

#Get the starting row from each stock excel file
def checkStartRow(stock):
    excelSheet = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Full_Data.xls'
    workbookOpen = xlrd.open_workbook(excelSheet)
    sheets = workbookOpen.sheets()
    rows = sheets[0].nrows #Number of rows in the excel sheet
    return rows

INSTALLER.main()

fileText = input("Enter the file with stock symbols: ")
fileText = fileText+'.txt'

x = True #Infinate loop
while x:
    #Get the time in seconds
    second = int(datetime.datetime.now().strftime('%S'))
    minute = int(datetime.datetime.now().strftime('%M'))
    hour = int(datetime.datetime.now().strftime('%H'))
    todayTime = second + (minute*60) + (hour*60*60) #Todays time in seconds
    inMarket = True #Boolean to hold T/F for stock market in time
    if not (todayTime >= 30600 and todayTime < 54000):
        print('Time is outside Stock Market Availability. Pausing for 10 minutes...')
        time.sleep(600) #Pause the script for 10 minutes

    #If the time is within the stock market operating hours, continue
    if(todayTime >= 30600 and todayTime < 54000):
        stocks = [] #Array to hold stock symbols
        counter = 0 #Counter to indicate when to update the 'startRow' varaible
        startRowNumber = [] #Array to hold the excel file starting row number for each stock
        
        numberOfStocks = 0 #Counter for the number of stocks
        #Open text file holding stock symbols
        #Add stock to 'stocks' array, and add 1 to counter
        file = open(fileText, 'r')
        for line in file:
            stocks.append(line.replace('\n', '')) #Add to array
            numberOfStocks += 1 #Increment counter
        file.close()

        while inMarket:      
            second2 = int(datetime.datetime.now().strftime('%S'))
            minute2 = int(datetime.datetime.now().strftime('%M'))
            hour2 = int(datetime.datetime.now().strftime('%H'))
            todayTime2 = second2 + (minute2*60) + (hour2*60*60)
            print(datetime.datetime.now().strftime('It is | %H:%M:%S'))
            for i in range(numberOfStocks):
                #Check for excel file existance
                checkExcelFile(stocks[i])
                
                #If 'counter' = 0, update the startRow variable
                if counter == 0:
                    startRowNumber.append(checkStartRow(stocks[i]))
                
                checkTextFilePath(stocks[i])
                service = callStockGrab(stocks[i])
                if(service == True and todayTime2 < 54000 and todayTime2 > 30600):
                    print('The market is unavailable, pausing for 5 minutes...')
                    time.sleep(300)
            time.sleep(10) #Pause the script for 10 seconds
            
            if(todayTime2 > 54000):
                inMarket = False
                break

            counter +=1 #Itterate the counter

        print('Finalized Data')
        print()
        for i in range(numberOfStocks):
            removeTextFilePath(stocks[i])
            Excel_Daily_Updater.start(stocks[i], startRowNumber[i])
            
        print('The market is now out.')






















                
