import selenium
import os
import shutil
import pandas as pd
import time
import win32com.client
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.action_chains import ActionChains



df = pd.read_excel("input_reports.xlsx", sheet_name="Details")

base_path = os.path.abspath(os.path.dirname(__file__))
PATH = os.path.join(base_path, 'chromedriver.exe')

# PATH = "chromedriver.exe"


def month_check_from(month_from):
    month_name = driver.find_element(By.XPATH, '//span[@class="ui-datepicker-month"]')
    for i in range(1,15):
        if month_name.text != month_from:
            next_month_button = driver.find_element(By.XPATH,'//a[@title="Prev"]')
            next_month_button.click()
            month_name = driver.find_element(By.XPATH, '//span[@class="ui-datepicker-month"]')
        else:
            break



def month_check_to(month_to):
    month_name = driver.find_element(By.XPATH, '//span[@class="ui-datepicker-month"]')
    for i in range(1,15):
        if month_name.text != month_to:
            next_month_button = driver.find_element(By.XPATH,'//a[@title="Prev"]')
            next_month_button.click()
            month_name = driver.find_element(By.XPATH, '//span[@class="ui-datepicker-month"]')
        else:
            break
    


def making_folder():
    reports_dir = 'Reports'
    if os.path.exists(reports_dir):
        shutil.rmtree(reports_dir)
    os.mkdir(reports_dir)

    for index, row in df.iterrows():
        module_alias = row['Module']
        client_alias = row['ClientAlias']
        client_dir = os.path.join(reports_dir, client_alias)
        if os.path.exists(client_dir):
            continue    
        else:
            os.mkdir(client_dir)
    
    
    for index, row in df.iterrows():
        module_alias = row['Module']
        client_alias = row['ClientAlias']
        module_dir = os.path.join(reports_dir, client_alias, module_alias)
        if os.path.exists(module_dir):
            continue    
        else:
            os.mkdir(module_dir)

# def renamingwithexceltitle():
#     reports_folder = os.path.abspath('Reports')
#     excel_app = win32com.client.Dispatch('Excel.Application')
#     # Loop through the subfolders in the Reports folder
#     for client_folder in os.listdir(reports_folder):
#         client_folder_path = os.path.join(reports_folder, client_folder)
#         if not os.path.isdir(client_folder_path):
#             continue
        
#         for module_folder in os.listdir(client_folder_path):
#             module_folder_path = os.path.join(client_folder_path, module_folder)
#             if not os.path.isdir(module_folder_path):
#                 continue
        
#             # Loop through the Excel files in the module folder
#             for excel_file_name in os.listdir(module_folder_path):
#                 excel_file_path = os.path.join(module_folder_path, excel_file_name)
#                 if not excel_file_name.endswith('.xlsx'):
#                     continue
                
#                 # Open the Excel file and get the subject property
                
#                 excel_app.Visible = False
#                 excel_workbook = excel_app.Workbooks.Open(excel_file_path)
#                 subject = excel_workbook.BuiltinDocumentProperties('Subject').Value
                
#                 # Rename the Excel file with the subject property
#                 new_file_name = f"{subject}.xlsx"
#                 new_file_path = os.path.join(client_folder_path, new_file_name)
#                 excel_workbook.Close(False)

#                 os.rename(excel_file_path, new_file_path)
                
#     excel_app.Quit()

#     dir_path = '/path/to/directory'

#     # Get a list of all the files in the directory
#     files = os.listdir(dir_path)

#     # Get the modification time of each file in the directory
#     mod_times = [os.path.getmtime(os.path.join(dir_path, f)) for f in files]

#     # Find the maximum modification time in the list
#     max_mod_time = max(mod_times)

#     # Get the path of the file with the maximum modification time
#     latest_file_path = os.path.join(dir_path, files[mod_times.index(max_mod_time)])

#     # Rename the latest file
#     os.rename(latest_file_path, os.path.join(dir_path, 'new_file_name.ext'))            


def login(i) :
    
    client_name = df.loc[i, "ClientAlias"]
    username = df.loc[i, "Username"]
    password = df.loc[i, "Password"]

    

    url = "https://www.ninjacrm.com/"+client_name+"/"

    driver.get(url)

    uid_ninja = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH,"/html/body/div[1]/div[4]/form/input[1]"))
        )
    uid_ninja.click()
    uid_ninja.send_keys(username)
        
        
    passwordofninjauploader = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/input[2]")
    passwordofninjauploader.send_keys(password)

    submit_button = driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/form/input[4]")
    submit_button.click()


def report_fetch(i):
    
    client_name = df.loc[i, "ClientAlias"]
    module = df.loc[i, "Module"]
    report_name = df.loc[i, "Report Name"]

    date_from = row['Date From'].strftime("%d").lstrip('0')
    month_from = row['Date From'].strftime("%B")
    year_from = row['Date From'].strftime("%Y")

    date_to = row['Date To'].strftime("%d").lstrip('0')
    month_to = row['Date To'].strftime("%B")
    year_to = row['Date To'].strftime("%Y")
    
   
    
    
    #opening reports in case something else has opened up(NinjaCRM ke backend team ko chain nhi milta  kyunki)

    url_reports = "https://www.ninjacrm.com/"+client_name+"/reports/"
    driver.get(url_reports)

    #selecting Insurance and Service as a test
    print('//span[@ng-bind="report.name" and text() ="' + module + '"]')
    module_select = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH,'//span[@ng-bind="report.name" and text() ="' + module + '"]'))
        )
    module_select.click()


    #skimming  through all the reports

    report_select = driver.find_element(By.XPATH,'//div[@class="report_title ng-binding" and text() ="' + report_name + '"]')
    report_select.click()

    time.sleep(5)
    try :
        date_from_popup = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.ID,"date_from"))
            )
        date_from_popup.click()
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_from_popup)
        month_check_from(month_from)
        year_drop = driver.find_element(By.XPATH,'//select[@class="ui-datepicker-year"]')
        year_drop.click()
        year_select = driver.find_element(By.XPATH,'//option[@value="'+year_from+'" and text()="'+year_from+'"]')
        year_select.click()
        date_select = driver.find_element(By.XPATH,'//a[@aria-current="false" and @data-date="'+date_from+'" and text()="'+date_from+'"]')
        date_select.click()

        date_to_popup = driver.find_element(By.ID,"date_to")
        driver.execute_script("arguments[0].removeAttribute('readonly')", date_to_popup)
        date_to_popup.click()
        month_check_to(month_to)
        year_drop = driver.find_element(By.XPATH,'//select[@class="ui-datepicker-year"]')
        year_drop.click()
        year_select = driver.find_element(By.XPATH,'//option[@value="'+year_to+'" and text()="'+year_to+'"]')
        year_select.click()
        date_select = driver.find_element(By.XPATH,'//a[@aria-current="false" and @data-date="'+date_to+'" and text()="'+date_to+'"]')
        date_select.click()
    except :
        print("Datebox not found")
        
    try: 
        select_all = driver.find_element(By.ID, "selectallcheckbox")
        select_all.click()
    except :
        
        print("Checkbox not found")

    download = driver.find_element(By.ID,"download_btn")

    download.click()

    WebDriverWait(driver, 1800).until(EC.element_to_be_clickable(download))
    WebDriverWait(driver, 30).until(lambda driver: driver.execute_script('return document.readyState') == 'complete')
    time.sleep(2)
    url_signout = "https://www.ninjacrm.com/"+client_name+"/index.php?logout=true"
    driver.get(url_signout)


#-------------------------------------code-------------------------------------------------------------------------------

op = webdriver.ChromeOptions()

making_folder()

excel_app = win32com.client.Dispatch('Excel.Application')
excel_app.Visible = False

for i, row in df.iterrows():
    client_alias = row['ClientAlias']
    module_alias = row['Module']
    fromdate = row['Date From'].strftime("%B %d, %Y")
    todate = row['Date To'].strftime("%B %d, %Y")

    reports_dir = os.path.join(base_path, 'Reports', client_alias, module_alias) 
    p = {'download.default_directory': reports_dir}
    op.add_experimental_option('prefs', p)
    driver = webdriver.Chrome(executable_path=PATH, options=op)
    driver.maximize_window()
    
    login(i)
    report_fetch(i)
    
    driver.quit()

    files = os.listdir(reports_dir)
    mod_times = [os.path.getmtime(os.path.join(reports_dir, f)) for f in files]
    max_mod_time = max(mod_times)
    latest_file_path = os.path.join(reports_dir, files[mod_times.index(max_mod_time)])
    excel_workbook = excel_app.Workbooks.Open(latest_file_path)
    
    try:
        subject = excel_workbook.BuiltinDocumentProperties('Subject').Value
        new_file_name = f"{subject}"
        excel_workbook.Close(False)
        os.rename(latest_file_path, os.path.join(reports_dir, new_file_name + '_' +fromdate+ ' to ' + todate +'.xlsx'))
    except:
        name = excel_workbook.Name
        new_file_name = f"{name}"
        excel_workbook.Close(False)
        os.rename(latest_file_path, os.path.join(reports_dir, new_file_name + '_' +fromdate+ ' to ' + todate +'.xlsx'))
    
                

excel_app.Quit()
