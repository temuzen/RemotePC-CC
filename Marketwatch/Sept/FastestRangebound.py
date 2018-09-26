#Nifty50MW.xlsx
#Nifty50MW.xlsx
#Optionswatch.xlsx

from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time
#App = xw.App()

sht = xw.Book(r'C:\Users\admin\Documents\GitHub\RemotePC-CC\Marketwatch\Sept\currencies.xlsx').sheets[0]
SAVED = xw.Book(r'C:\Users\admin\Documents\GitHub\RemotePC-CC\Marketwatch\Sept\Book_26_9_data.xlsx').sheets[0]

#wrt = xw.books('Book1').sheets[0]
wrt = SAVED
i = 1
	

while(1):    
    s1 = time.time()
    String1 = str(sht.range('A2:M110').value)
    e1 = time.time()
   # print("Time S1:"+str(e1-s1))
    s2x = time.time()
    s2 = time.time()
    String2 = str(sht.range('A2:M110').value)
    e2 = time.time()
  #  print("Time S2:"+str(e2-s2))
    #String1 = String1 + str(sht.range('A'+str(x)+':'+'M'+str(x)).value)
    
  #  print(len(String1))
  #  print(String1)
  #  print(s2-s2x)
    if String1==String2:
        continue    
    else:
        wrt.range('A'+str(i+1)).value = String2          
        i=i+1
   