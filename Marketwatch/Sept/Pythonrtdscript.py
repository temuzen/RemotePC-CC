#Nifty50MW.xlsx
#Nifty50MW.xlsx
#Optionswatch.xlsx

from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time
#App = xw.App()

sht = xw.Book(r'D:\Codes\Github\Temuzen\RemotePC-CC\Marketwatch\Sept\Optionswatch.xlsx')
wrt = xw.books('book1').sheets[2]
i = 1

while(1):    

    for x in range(2,100):         
        String1= '' 
        String1 = String1 + str(sht.range('A'+str(x)+':'+'M'+str(x)).value)

    for y in range(2,100):         
        String2 = '' 
        String2 = String2 + str(sht.range('A'+str(y)+':'+'M'+str(y)).value)

    if String1==String2:
        continue    
    else:
        wrt.range('A'+str(i+1)).value = String2          
        i=i+1
