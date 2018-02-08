import json
import requests
import os
import os.path
import xlwt
import xlrd
from xlutils.copy import copy
import datetime

def grabStock(stock):
    #Get the link
    rsp = requests.get('https://finance.google.com/finance?q=' + stock + '&output=json')
    if(rsp.reason == 'Service Unavailable'):
         return True #The json service is no longer available
                     #Retuning True to pause the script for 10 minutes
    if rsp.status_code in (200,):

        #File location for the daily changing excel sheet
        excelSheet = 'Excel_Files/'+stock.upper()+'/Excel_Sheet_'+stock.upper()+'_Full_Data.xls'
        book = xlwt.Workbook()

        #Get the number of rows currently used in the excel sheet
        #Used for placing new data
        workbookOpen = xlrd.open_workbook(excelSheet)
        sheets = workbookOpen.sheets()
        rows = sheets[0].nrows #Number of rows in the excel sheet
        
        last = 0.0 #The last price placed in the excel sheet
        stockRange = 0.0 #The range between the last price and current price

        fin_data = json.loads(rsp.content[6:-2].decode('unicode_escape')) #Gathering the data from the web page
        hold = fin_data['l'] #The current price of the stock
        StockCompany = fin_data['name'] #The company name
        priceDifference = 0.0 #The value used to mark when to add new data to the excel file
        if(float(hold) > 75.00): #If above $75, add data when the last record and current price are $0.10 apart
            priceDifference = 0.10
        elif(float(hold) > 35.00):
            priceDifference = 0.06
        else:
            priceDifference = 0.02

        #To add another price range slot it in where it should be and fill in this (remove the pound symbols):
        #elif(float(hold) > stockPrice):
        #   priceDifference = priceDifference

        lastInformation = 'lastInformation_'+stock.upper()+'.txt' #File used to hold information from the last iteration of the script

        #Gather data from the lastInformation text file
        with open(lastInformation, 'r') as fp:
            for line in fp:
                last = float(line.replace('\n', ''))
        fp.close()
        
        stockRange = abs(last - float(hold)) #Get the price difference from the last and current recordings. Compared with the hard coded price differences
        now = datetime.datetime.now() #Get the date
        
        #Print out basic current information
        print('Company: ' + StockCompany)
        print('Last: ' + str(last) + ' - ' + str(hold))
        print('Stock Range: ' + str(stockRange))
        print()

        #If the price difference that was hard coded above is less than the stock range, add this information to the excel file
        #Otherwise skip, and run again in 10 seconds
        if(priceDifference < stockRange):

            #get sentiment analysis information
            
            rb = xlrd.open_workbook(excelSheet) #Open the excel file
            sh = copy(rb) #Copy the file to a writable doc
            sh_sheet = sh.get_sheet(0) #Use sheet 0
            #Write to the database on the next available row
            sh_sheet.write(rows,0,now.strftime("%m/%d/%Y | %H:%M"))

            #Get and set the time in seconds
            now = datetime.datetime.now()
            second = int(now.strftime('%S'))
            minute = int(now.strftime('%M'))
            hour = int(now.strftime('%H'))
            todayTime = second + (minute*60) + (hour*60*60) #Todays time in seconds
            sh_sheet.write(rows,1,todayTime)
            
            sh_sheet.write(rows,2,float(format(fin_data['l'])))
            sh_sheet.write(rows,3,float(format(fin_data['c'])))
            sh_sheet.write(rows,4,float(format(fin_data['cp'])))
            sh_sheet.write(rows,5,float(format(fin_data['op'])))
            sh_sheet.write(rows,6,float(format(fin_data['lo'])))
            sh_sheet.write(rows,7,float(format(fin_data['hi'])))

            #Get the volume traded today, and change to integer if string
            #If in millions (m), multiply value by 10^6
            #If in billions (b), multiply value by 10^9
            tradedChange = 0;
            if 'M' in format(fin_data['vo']):
                tradedChange = float(format(fin_data['vo'].replace('M', ''))) * 10**6
            if 'B' in format(fin_data['vo']):
                tradedChange = float(format(fin_data['vo']).replace('B', '')) * 10**9
            sh_sheet.write(rows,8,tradedChange)

            #Get the volume traded average, and change to integer if string
            #If in millions (m), multiply value by 10^6
            #If in billions (b), multiply value by 10^9
            tradedAverageChange = 0;
            if 'M' in format(fin_data['avvo']):
                tradedAverageChange = float(format(fin_data['avvo']).replace('M', '')) * 10**6
            if 'B' in format(fin_data['avvo']):
                tradedAverageChange = float(format(fin_data['avvo']).replace('B', '')) * 10**9
            sh_sheet.write(rows,9,tradedAverageChange)
            
            sh_sheet.write(rows,10,float(format(fin_data['lo52'])))
            sh_sheet.write(rows,11,float(format(fin_data['hi52'])))

            #Change Price/Earnings Ratio to a float
            if '.' in format(fin_data['pe']):
                peChange = float(format(fin_data['pe']))
            else:
                peChange = '-'
            sh_sheet.write(rows,12,peChange)
            
            sh_sheet.write(rows,13,float(format(fin_data['eps'])))

            if '.' in format(fin_data['dy']):
                dyChange = float(format(fin_data['dy']))
            else:
                dyChange = '-'
            sh_sheet.write(rows,14,dyChange)

            if '.' in format(fin_data['ldiv']):
                ldivChange = float(format(fin_data['ldiv']))
            else:
                ldivChange = '-'
            sh_sheet.write(rows,15,ldivChange)

            #Get the total number of shares available, and change to integer if string
            #If in millions (m), multiply value by 10^6
            #If in billions (b), multiply value by 10^9
            sharesChange = 0;
            if 'M' in format(fin_data['shares']):
                sharesChange = float(format(fin_data['shares']).replace('M', '')) * 10**6
            if 'B' in format(fin_data['shares']):
                sharesChange = float(format(fin_data['shares']).replace('B', '')) * 10**9
            sh_sheet.write(rows,16,sharesChange)

            #Get the Market Cap, and change to integer if string
            #If in millions (m), multiply value by 10^6
            #If in billions (b), multiply value by 10^9
            marketCapChange = 0;
            if 'M' in format(fin_data['mc']):
                sharesChange = float(format(fin_data['mc']).replace('M', '')) * 10**6
            if 'B' in format(fin_data['mc']):
                marketCapChange = float(format(fin_data['mc']).replace('B', '')) * 10**9
            sh_sheet.write(rows,17,marketCapChange)
            
            sh_sheet.write(rows,18,float(format(fin_data['beta'])))
            
            #Change the % to decimal (e.i 62% to 0.62)
            instOwnChange = 0
            instString = format(fin_data['instown']).replace('%', '')
            instOwnChange = float(instString) * 10**-2
            sh_sheet.write(rows,19,instOwnChange)
            
            #Format the columns to 20 characters wide
            for i in range(20):
                sh_sheet.col(i).width = 20*256
            sh.save(excelSheet) #Save the excel file

            #Print out the information gathered to the console
            print('Current Price: {}'.format(fin_data['l']))
            print('Change: {}'.format(fin_data['c']))
            print('Change Percentage: {}'.format(fin_data['cp']))
            print('Opening Price: {}'.format(fin_data['op']))
            print('Price/Earnings Ratio: {}'.format(fin_data['pe']))
            print('52-week high: {}'.format(fin_data['hi52']))
            print('52-week low: {}'.format(fin_data['lo52']))
            print('Outstanding Shares: {}'.format(fin_data['shares']))
            print('Highest Price: {}'.format(fin_data['hi']))
            print('Lowest Price: {}'.format(fin_data['lo']))
            print('Yeild: {}'.format(fin_data['dy']))
            print('Dividend: {}'.format(fin_data['ldiv']))
            print('Earnings Per Share: {}'.format(fin_data['eps']))
            print('Beta: {}'.format(fin_data['beta']))
            print('Institutional Ownership: {}'.format(fin_data['instown']))
            print('Market Cap: {}'.format(fin_data['mc']))
            print('Daily Traded Average: {}'.format(fin_data['avvo']))
            print('Traded Today: {}'.format(fin_data['vo']))
            print()

            #Update the lastInformation text file with new data
            file2 = open(lastInformation, 'w+')
            file2.write(hold)
            file2.close()
        print('-------------------------------------------------------------')
        print()
        #Check back with Start Up script to see if restarting
        return False #The json service is still available
