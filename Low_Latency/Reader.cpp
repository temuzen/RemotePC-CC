#include<iostream>
#include<fstream>
using namespace std;

int main(){
    char data[1000];
    //ofstream Myexcelfile;
    fstream Myexcelfile; 
    Myexcelfile.open("D:\Codes\Github\Temuzen\RemotePC-CC\Low_Latency\Commodities.xlsx");
    cout << "Reading from the file" << endl; 
    Myexcelfile >> data;

}



