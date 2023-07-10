
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook
import os
import glob



folder_location = "C:\\Users\\Jarvis\\Desktop\\IyunGelen"
excel_path = 'C:\\Users\\Jarvis\\Desktop\\resulttest\\RIyunGelen.xlsx'

excel =  load_workbook(filename=excel_path)
forLenthexcel = pd.read_excel(excel_path,sheet_name='result')

length = forLenthexcel.shape[0]+1
activeSheet = excel['result']
i=1
files = glob.glob(f'{folder_location}\\*')

print(type(files))

while i <= length:
    inexcelFilePath = activeSheet[f"C{i}"].value
    inexcelFileName = str(inexcelFilePath).split(f"\\")[-1]
    indexOfFile = files.index(f"{folder_location}\\{inexcelFileName}")
    newName = activeSheet[f"A{i}"].value
    val1 = str(files[indexOfFile])
    val2 = f'C:\\Users\\Jarvis\\\\Desktop\\IyunGelen\\{newName}.xhtml'
    os.rename(val1,val2)
    print(f"{inexcelFileName} ->{newName}")
    i += 1
 

