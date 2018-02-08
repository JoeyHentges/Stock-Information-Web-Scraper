# Stock-Information-Web-Scraper
Using the Python language, and Google Finance, gather stock information throughout the day, and log it in an Excel Sheet.

To start the program open the 'Initializer' python file. Any packages that need to be installed, will be installed (using the 'INSTALLER' python file).

All data is stored in the folder titled 'Excel_Files', and each stock has their own folder holding their two excel files.  

To view, change, or add stocks that you want to gather the information on, open and edit the 'Stock_Symbols.txt' file. Keep the formating the same; a single stock symbol (capitalized) on each line. To run the scipt more efficiently create multiple text files to hold the stock symbols. It currently takes between 5-10 seconds to run through 14 stocks.  

In the program, under the 'Stock_Price_Grab.py' file, there is a group of 'If' statements which determine when the stock logs new information, i.e. when a stocks last logged price has changed by a given ammount. For example, if you wish to change how often (the range) a stock $75.00 or more is logged, it checks if the stock's current price, and last logged price have a price range of $0.10. If it does it logs the data, otherwise it does not. Change either the stocks price, or range of price to alter when it logs the data.  

If you wish to add another price range, in the same python file, use this outline, and slot the new else if statement into the script. Slot it in above the next lowest stock price, and below the next largest stock price. If it is the largest change the 'elif' to 'if', and change the current 'if' to 'elif'. An example is comment listed inside the 'Stock_Price_Grab.py' file.  

Other than making sure to have the Python Language installed, and any packages that don't correctly install, this script should run indefinately as long as the program isn't closed.
