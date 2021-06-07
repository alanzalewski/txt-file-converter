'''This program opens a window that allows the user to select txt files and convert 
them to either xlsx or csv format, outputting them to the same directory.'''

import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

root = tk.Tk()

canvas1 = tk.Canvas(root, width=250, height=110, bg='white')
canvas1.pack()

def convertToXlsx():
    # allows user to select txt files for conversion
    import_file_path = filedialog.askopenfilenames()
    
    # loops through each selected file
    for entry in import_file_path:
        export_file_path = entry.replace('.txt','.xlsx')
    
        with open(entry,'r') as txtfile:
	    # formatting data
            rawdata = [row.split('\t') for row in txtfile]
            r=0
            for row in rawdata:
                if r<len(rawdata)-1:
                    row[-1] = row[-1][0:-1]
                r+=1
        
            # writing to xlsx file
            df = pd.DataFrame(rawdata)
            df.loc[7:len(rawdata),0:8] = df.loc[7:len(rawdata),0:8].astype(float)
            excelfile = pd.ExcelWriter(export_file_path, engine='xlsxwriter')
            df.to_excel(excelfile, sheet_name='Sheet1', index=False, header=False)
            excelfile.save()
            excelfile.close()
    
def convertToCsv():
    # allows user to select txt files for conversion
    import_file_path = filedialog.askopenfilenames()
    
    # loops through each selected file
    for entry in import_file_path:
        export_file_path = entry.replace('.txt','.csv')
        
        with open(entry,'r') as txtfile:
	    # formatting data
            rawdata = [row.split('\t') for row in txtfile]
            r=0
            for row in rawdata:
                if r<len(rawdata)-1:
                    row[-1] = row[-1][0:-1]
                r+=1
            
            # writing to csv file
            df = pd.DataFrame(rawdata)
            df.loc[7:len(rawdata),0:8] = df.loc[7:len(rawdata),0:8].astype(float)
            df.to_csv(export_file_path, index=False, header=False)
    
saveAsButtonXlsx = tk.Button(text='Select files to convert to xlsx', command=convertToXlsx, bg='gainsboro', fg='black', font=('helvetica',10))
canvas1.create_window(125,35,window=saveAsButtonXlsx)

saveAsButtonCsv = tk.Button(text='Select files to convert to csv', command=convertToCsv, bg='gainsboro', fg='black', font=('helvetica',10))
canvas1.create_window(125,80,window=saveAsButtonCsv)

root.mainloop()
