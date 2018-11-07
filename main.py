#Use 
# C:\Anaconda3\python.exe  C:\Anaconda3\cwp.py C:\Anaconda3 C:\Anaconda3\python.exe main.py
#  to run so it loads anaconda dependencies

import time
import win32com.client as win32

import pandas as pd

def excel():
    """"""
    xl = win32.gencache.EnsureDispatch('Excel.Application')
    ss = xl.Workbooks.Add()
    sh = ss.ActiveSheet
 
    xl.Visible = True
    time.sleep(1)
 
    sh.Cells(1,1).Value = 'Hacking Excel with Python Demo'
 
    time.sleep(1)
    for i in range(2,8):
        sh.Cells(i,1).Value = 'Line %i' % i
        time.sleep(1)
 
    ss.Close(False)
    xl.Application.Quit()
 
if __name__ == "__main__":
    pass
    # excel()