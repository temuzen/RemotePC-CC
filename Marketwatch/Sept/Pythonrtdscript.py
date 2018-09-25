#Nifty50MW.xlsx
#Nifty50MW.xlsx
#Optionswatch.xlsx

from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time
App = xw.App()

sht = xw.Book(r'C:\Users\admin\Documents\GitHub\RemotePC-CC\Marketwatch\Sept\watchlist.xlsx').sheets[0]
wrt = xw.books('Book1').sheets[0]
i = 1
	

wrt.range('A1').value = 'Trading symbol'
wrt.range('B1').value = 'LTP'
wrt.range('C1').value = 'Bid qty'
wrt.range('D1').value = 'Bid rate'
wrt.range('E1').value = 'Ask rate'
wrt.range('F1').value = 'Ask qty'
wrt.range('G1').value = 'LTQ'
wrt.range('H1').value = 'Volume traded today'
wrt.range('I1').value = 'Open interest'
wrt.range('J1').value = 'Total bid qty'
wrt.range('K1').value = 'Total ask qty'
wrt.range('L1').value = 'LTT'
wrt.range('M1').value = 'LUT'


while(1):    

    Str1timeS = time.time()
    for x in range(2,110):         
        String1= '' 
        String1 = String1 + str(sht.range('A'+str(x)+':'+'M'+str(x)).value)
    Str1timeE = time.time()
    print("Str1 time: " + str(Str1timeE-Str1timeS))
  
    Str2timeS = time.time()
    for y in range(2,110):         
        String2 = '' 
        String2 = String2 + str(sht.range('A'+str(y)+':'+'M'+str(y)).value)
    Str2timeE = time.time()
    print("Str2 time: " +str(Str2timeE-Str2timeS))
    
    CompS = time.time()  
    if String1==String2:
        continue    
    else:
        wrt.range('A'+str(i+1)).value = String2          
        i=i+1
    CompE = time.time()  

    print("Comp time: " +str(CompE-CompS))
