
from tkinter import filedialog
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook
import os
import glob
import tkinter as tk

root = tk.Tk()
root.title("Name Changer")
root.geometry("300x300")

def browse_FolderButton_click():
    folder_loc = filedialog.askdirectory()
    folder_loc_entry.insert(tk.END, folder_loc)  

def browse_FileButton_click():
    excel_loc = filedialog.askopenfile().name
    excel_loc_entry.insert(tk.END, excel_loc)  

def FileNameChanger(folder_loc, excel_loc):
    # folder_location = "C:\\Users\\Jarvis\\Desktop\\IyunGelen"
    # excel_path = 'C:\\Users\\Jarvis\\Desktop\\resulttest\\RIyunGelen.xlsx'

    folder_location = str(folder_loc).replace("/","\\")
    excel_path = str(excel_loc).replace("/","\\")

    excel =  load_workbook(filename=excel_path)
    forLenthexcel = pd.read_excel(excel_path,sheet_name='result')

    length = forLenthexcel.shape[0]+1
    activeSheet = excel['result']
    files = glob.glob(f'{folder_location}\\*')

    print(type(files))

    i=1

    while i <= length:
        inexcelFilePath = activeSheet[f"C{i}"].value
        inexcelFileName = str(inexcelFilePath).split(f"\\")[-1]
        indexOfFile = files.index(f"{folder_location}\\{inexcelFileName}")
        newName = activeSheet[f"A{i}"].value
        val1 = str(files[indexOfFile])
        val2 = f'{folder_location}\\{newName}.xhtml'
        os.rename(val1,val2)
        print(f"{inexcelFileName} ->{newName}")
        i += 1
 

def start_button_click():
    folder_loc = folder_loc_entry.get()
    excel_loc = excel_loc_entry.get()
    
    print("File path:", folder_loc)
    print("Resoruce File:", excel_loc)
    
    FileNameChanger(folder_loc, excel_loc)

folder_loc_label = tk.Label(root, text="download File Location:")
folder_loc_label.pack()
folder_loc_entry = tk.Entry(root)
folder_loc_entry.pack()

browse_button = tk.Button(root, text="Browse", command=browse_FolderButton_click)
browse_button.pack()

excel_loc_label = tk.Label(root, text="Resource File Location:")
excel_loc_label.pack()
excel_loc_entry = tk.Entry(root)
excel_loc_entry.pack()

browse_button_forFile = tk.Button(root, text="Browse", command=browse_FileButton_click)
browse_button_forFile.pack()

start_button = tk.Button(root, text="Start", command=start_button_click)
start_button.pack()

root.mainloop()

