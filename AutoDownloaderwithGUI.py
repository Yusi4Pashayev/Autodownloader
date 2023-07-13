import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import threading
import os
import glob

chrome_options = Options()

#prosesi dayandirmaq ucun
run = True

#e-qaimeden etibarli cixis
def exit(browser,wait2):
    try:
        browser.find_element(By.CSS_SELECTOR,'.icon-switch2').click()
    except:
        print("exiting")
    try:
        wait2.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.sa-button-container>.cancel'))) 
        browser.find_element(By.CSS_SELECTOR,'.sa-button-container>.cancel').click()
    except:
        print("exiting")

#son yuklenen mhtml in adini geri qaytarir
def DownloadFileName(folder_location):
    files = glob.glob(f'{folder_location}\\*')
    files.sort(key=os.path.getmtime)
    file = files[-1]
    print(file)
    return file
    
# def ChangeFileName(excelResult):
#     length = excelResult.shape[0]
#     start
#     while 
#     print("change file name")  


def StartDownload(username, password, pin, optionValue, result_file_name, start_number, file_path,downloadfile_path):
#sisteme giris
    adress = 'https://login.e-taxes.gov.az/login/'
    username = username
    password_2 = password
    password_1 = pin
    optionV = optionValue
    newResultFileName = result_file_name
    startNumber = int(start_number)-1
    filePath = file_path

#default download folderini deyisir
    new_downloadfile_path = str(downloadfile_path).replace("/","\\")
    chrome_options.add_experimental_option("prefs", {"download.default_directory": f'{new_downloadfile_path}'})

#browseri aktivlesdirir
    browser = webdriver.Chrome(options=chrome_options)

#drowserde 5 saniye gozleme muddeti yaradir
    wait2 = WebDriverWait(browser, 5)

#address linkine gedir
    browser.get(adress)

#sisteme daxil olur
    login = browser.find_element(By.CSS_SELECTOR,'[src="images/icons/user.png"]')
    login.click()
    wait2.until(EC.visibility_of_element_located((By.ID, 'username')))
    username_field = browser.find_element(By.ID,'username')
    username_field.send_keys(username)
    password2 = browser.find_element(By.ID,'password2')
    password2.send_keys(password_2)
    password1 = browser.find_element(By.ID,'password1')
    password1.send_keys(password_1)
    select = browser.find_element(By.ID,'idare')
    select.send_keys('E-qaimə')
    browser.find_element(By.CSS_SELECTOR,'#Section2 button').click()

# mhtml file i yukleyir
    #double click ucun
    actionChains = ActionChains(browser)   

    excel = pd.read_excel(f'{os.getcwd()}\\Qaime.xlsx',sheet_name=optionV)
    excelResult = load_workbook(filename=os.getcwd() + '\\Result.xlsx')
    sheet = excelResult.active
    b = excel.shape[0]
    browser.execute_script("window.open('');")
    browser.switch_to.window(browser.window_handles[1])

    while startNumber<b and run == True:
        link = excel.loc[startNumber,'link']
        qaimeNum = excel.loc[startNumber,'E-QAIME']
        
        condition = False
        while condition == False:
            try:
                browser.get(link)
                wait2.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#resultArea tbody td')))
                condition = True
            except:
                condition = False

        response = browser.find_elements(By.CSS_SELECTOR, '#resultArea tbody td')
        browser.implicitly_wait(2)
        actionChains.double_click(response[4]).perform()
  
        parts = browser.find_elements(By.CSS_SELECTOR,'.nav-tabs-bottom>li>a')
        try:
            wait2.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '.nav-tabs-bottom>li>a')))
            parts = browser.find_elements(By.CSS_SELECTOR,'.nav-tabs-bottom>li>a')
        except:
            actionChains.double_click(response[4]).perform()
            parts = browser.find_elements(By.CSS_SELECTOR,'.nav-tabs-bottom>li>a')
        if len(parts) == 0:
            sheet[f'A{startNumber+1}']= qaimeNum
            sheet[f'B{startNumber+1}']= f'{browser.current_url}'
            sheet[f'D{startNumber+1}'] = "sehife tapilmadi"
            excelResult.save(filename=f'{filePath}\\R{newResultFileName}.xlsx')
        else:
            try:
                parts[3].click()
                wait2.until(EC.visibility_of_element_located((By.CSS_SELECTOR, '#showdata button')))
                downloads = browser.find_elements(By.CSS_SELECTOR,'#showdata button')
                if len(downloads) == 0:
                    sheet[f'A{startNumber+1}']= qaimeNum
                    sheet[f'B{startNumber+1}']= f'{browser.current_url}'
                    sheet[f'D{startNumber+1}'] = "yuklenmedi"
                    excelResult.save(filename=f'{filePath}\\R{newResultFileName}.xlsx')
                else:
                    downloads[len(downloads)-1].click()
                    sheet[f'A{startNumber+1}']= qaimeNum
                    sheet[f'B{startNumber+1}']= f'{browser.current_url}'
                    excelResult.save(filename=f'{filePath}\\R{newResultFileName}.xlsx')
                    sheet[f'C{startNumber+1}']= f'{DownloadFileName(new_downloadfile_path)}'
                    sheet[f'D{startNumber+1}'] = "yuklendi"
                    # if startNumber == b-1:
                    #     # ChangeFileName(excelResult)
            except:
                sheet[f'A{startNumber+1}']= qaimeNum
                sheet[f'B{startNumber+1}']= f'{browser.current_url}'
                sheet[f'D{startNumber+1}'] = "ELEMENT TAPILMADI"
                excelResult.save(filename=f'{filePath}\\R{newResultFileName}.xlsx')
                continue
        if startNumber == b-1:
            exit(browser,wait2)

        startNumber += 1

def start_button_click():
    username = username_entry.get()
    password = password_entry.get()
    pin = pin_entry.get()
    optionValue = "{}".format(value_inside.get())
    result_file_name = result_file_entry.get()
    start_number = start_number_entry.get()
    downloadfile_path = downloadfile_path_entry.get()
    file_path = file_path_entry.get()

    print("Username:", username)
    print("Password:", password)
    print("PIN:", pin)
    print(optionValue)
    print("Result File Name:", result_file_name)
    print("Start Number:", start_number)
    print("download File path:", downloadfile_path)
    print("File Path:", file_path)
    StartDownload(username, password, pin, optionValue, result_file_name, start_number, file_path, downloadfile_path)

def browse_button_click():
    file_path = filedialog.askdirectory()
    file_path_entry.insert(tk.END, file_path)  

def browse_button_dclick():
    downloadfile_path = filedialog.askdirectory()
    downloadfile_path_entry.insert(tk.END, downloadfile_path)  

def stop_button_click():
    global run
    run =False
    exit()


root = tk.Tk()
root.title("qaime Formu")
root.geometry("400x600")

username_label = tk.Label(root, text="İstifadəçi kodu:")
username_label.pack()
username_entry = tk.Entry(root)
username_entry.pack()

password_label = tk.Label(root, text="Parol:")
password_label.pack()
password_entry = tk.Entry(root, show="*")
password_entry.pack()

pin_label = tk.Label(root, text="Şifrə:")
pin_label.pack()
pin_entry = tk.Entry(root)
pin_entry.pack()

optionsList = ["gelenler", "gonderilenler"]
value_inside = tk.StringVar(root)
value_inside.set("seçim edin")
selectValue = tk.OptionMenu(root,value_inside,*optionsList)
selectValue.pack()

result_file_label = tk.Label(root, text="Result File Name:")
result_file_label.pack()
result_file_entry = tk.Entry(root)
result_file_entry.pack()

start_number_label = tk.Label(root, text="Start Number:")
start_number_label.pack()
start_number_entry = tk.Entry(root)
start_number_entry.pack()

downloadfile_path_label = tk.Label(root, text="download File Location:")
downloadfile_path_label.pack()
downloadfile_path_entry = tk.Entry(root)
downloadfile_path_entry.pack()

browse_dbutton = tk.Button(root, text="Browse", command=browse_button_dclick)
browse_dbutton.pack()

file_path_label = tk.Label(root, text="File Path:")
file_path_label.pack()
file_path_entry = tk.Entry(root)
file_path_entry.pack()

browse_button = tk.Button(root, text="Browse", command=browse_button_click)
browse_button.pack()

startThread = threading.Thread(target=start_button_click)
stopThread = threading.Thread(target=stop_button_click)

start_button = tk.Button(root, text="Start", command=startThread.start)
start_button.pack()

stop_button = tk.Button(root, text="Stop", command=stopThread.start)
stop_button.pack()

root.mainloop()




