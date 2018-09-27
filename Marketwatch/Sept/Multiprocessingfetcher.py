from xlwings import Book, Sheet, Range, Chart
import threading
from multiprocessing import Process
import xlwings as xw 
import time



from xlwings import Book, Sheet, Range, Chart
import xlwings as xw
import time

def Writer_function(inputfile,size,outputfile):

    sht = xw.Book(r'C:\\Users\\Yash\Documents\Datafeeds\\'+inputfile).sheets[0]
    Sheet = xw.Book(r'C:\\Users\\Yash\Documents\Datafeeds\\'+inputfile).sheets[1]
    wrt = Sheet
    txt = open(outputfile,'w')
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
      #  print("Time S1:"+str(e1-s1))
        s2x = time.time()
        s2 = time.time()
        String2 = str(sht.range('A2:M'+str(size)).value)
        e2 = time.time()
     #   print("Time S2:"+str(e2-s2))
        #String1 = String1 + str(sht.range('A'+str(x)+':'+'M'+str(x)).value)
        
    #  print(len(String1))
    #  print(String1)
      #  print(s2-s2x)
        if String1==String2:
            continue    
        else:
     #       wrt.range('A'+str(i+1)).value = String2          
            txt.writelines(String2)
            i=i+1
       #     print("Comp time: " +str(CompE-CompS))


if __name__=="__main__":
        
    p1 = Process(target=Writer_function, args=("Nifty50.xlsx",50))
    p2 = Process(target=Writer_function, args=("Currency.xlsx",26))
    p3 = Process(target=Writer_function,args=("BankNifty.xlsx",12))
    p4 = Process(target=Writer_function,args=("Options.xlsx",50))
    p5 = Process(target=Writer_function,args=("optionsnextexpiry.xlsx",51))
    p6 = Process(target=Writer_function,args=("Commodities.xlsx",51,"Outputfile.csv"))


    #p1.start()
    p6.start()
    #p3.start()
   # p4.start()
   # p5.start()
    