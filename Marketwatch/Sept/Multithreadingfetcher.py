from xlwings import Book, Sheet, Range, Chart
import threading
import xlwings as xw 
import time

"""
def Options():

def Nifty50():

def BankNifty():

def Currencies():
"""

#Nifty50MW.xlsx
#Nifty50MW.xlsx
#Optionswatch.xlsx

from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time

def Writer_function(inputfile,size):

    sht = xw.Book(r'C:\\Users\\admin\Documents\Datafeeds\\'+inputfile).sheets[0]
    Sheet = xw.Book(r'C:\\Users\\admin\Documents\Datafeeds\\'+inputfile).sheets[1]
    wrt = Sheet
    i = 1

    wrt.range('A1').value = sht.range('A1').value
    wrt.range('B1').value = sht.range('B1').value
    wrt.range('C1').value = sht.range('C1').value
    wrt.range('D1').value = sht.range('D1').value
    wrt.range('E1').value = sht.range('E1').value
    wrt.range('F1').value = sht.range('F1').value
    wrt.range('G1').value = sht.range('G1').value
    wrt.range('H1').value = sht.range('H1').value
    wrt.range('I1').value = sht.range('I1').value
    wrt.range('J1').value = sht.range('J1').value
    wrt.range('K1').value = sht.range('K1').value
    wrt.range('L1').value = sht.range('L1').value
    wrt.range('M1').value = sht.range('M1').value

    while(1):    
        s1 = time.time()
        String1 = str(sht.range('A2:M'+str(size)).value)
        e1 = time.time()
        print("Time S1:"+str(e1-s1))
        s2x = time.time()
        s2 = time.time()
        String2 = str(sht.range('A2:M'+str(size)).value)
        e2 = time.time()
        print("Time S2:"+str(e2-s2))
        #String1 = String1 + str(sht.range('A'+str(x)+':'+'M'+str(x)).value)
        
    #  print(len(String1))
    #  print(String1)
        print(s2-s2x)
        if String1==String2:
            continue    
        else:
            wrt.range('A'+str(i+1)).value = String2          
            i=i+1
            print("Comp time: " +str(CompE-CompS))

Writer_function("Nifty50.xlsx",100)

