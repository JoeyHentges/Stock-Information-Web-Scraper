#Script holding all of the function used in the Gather_Information.py script

#Function to get the average value from an array
def getAverage(array):
    arraySum = sum(array) #Value to hold the sum of the array
    return arraySum / len(array)

#Function to get the lowest value from an array
def getLowestPrice(array):
    return min(array)

#Function to get the time of the lowest price
#Get the time of its first occurance
def getLowestPriceTime(value, priceArray, timeArray):
    time = 0
    lowestPriceRow = 0
    #For loop to find the lowest price's row
    for i in range(len(priceArray)):
        if(value == priceArray[i]):
            lowestPriceRow = i+1
            print('Lowest price row: '+str(lowestPriceRow))
            break

    #Value from the array by the row number given
    time = timeArray[lowestPriceRow-1]

    return time, lowestPriceRow

#Function to get the volume when the lowest price occurs
def getLowestPriceVolume(row, volumeArray):
    return volumeArray[row-1]

#Function to get the lowest value from an array
def getHighestPrice(array):
    return max(array)

#Function to get the time of the highest price
#Get the time of its first occurance
def getHighestPriceTime(value, priceArray, timeArray):
    time = 0
    highestPriceRow = 0
    #For loop to find the highest price's row
    for i in range(len(priceArray)):
        if(value == priceArray[i]):
            highestPrice = i+1
            print('Highest price row: '+str(highestPriceRow))
            break

    #Value from the array by the row number given
    time = timeArray[highestPriceRow-1]

    return time, highestPriceRow

#Function to get the volume when the highest price occurs
def getHighestPriceVolume(row, volumeArray):
    return volumeArray[row-1]

#Function to get the 52 week low for the stock





