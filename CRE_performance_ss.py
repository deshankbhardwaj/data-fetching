import selenium
import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import filedialog
from pathlib import Path

PATH = "chromedriver.exe"
driver = webdriver.Chrome(PATH)
driver.maximize_window()

df = pd.read_excel("input_cre.xlsx")


def login(i) :
    
    client_name = df.loc[i, "ClientAlias"]
    username = df.loc[i, "Username"]
    password = df.loc[i, "Password"]
    module = df.loc[i, "Module"]

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


def ss(i) :
    
    client_name = df.loc[i, "ClientAlias"]
    username = df.loc[i, "Username"]
    password = df.loc[i, "Password"]
    module = df.loc[i, "Module"]
    

    url_creperf = "https://www.ninjacrm.com/"+client_name+"/new_cre_perf.php"
    driver.get(url_creperf)
    

    #time.sleep(3)

    iframe = driver.find_element(By.XPATH,"/html/body/iframe[1]")
    driver.switch_to.frame(iframe)

    module_drop = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[1]/div[1]/div[2]/i')
    module_drop.click()

    module_select = WebDriverWait(driver, 30).until(
        EC.presence_of_element_located((By.XPATH,"//div[@class='label' and text()='"+module+"']"))
        )

    
        #module_select = driver.find_element(By.XPATH,"//div[@class='label' and text()='"+module+"']")
                                        #'//*[@id="app"]/div/div[2]/div[1]/div[1]/div[3]/div/div['+module[-1]+']')
    module_select.click()
    
    time.sleep(3)

    day_picker = driver.find_element(
        By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[1]')
    day_picker.click()

    time.sleep(1)
    
    perf_table = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div[3]/div/div[2]/div[4]/div[2]/div/table"))
    )
    
    driver.execute_script('arguments[0].style.overflow = "visible";arguments[0].style.height = "832 px";', perf_table)

    date_time_ninja = driver.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/input')

    
    
    perf_table.screenshot(os.path.join(os.getcwd(), "screenshots", client_name + "_" + date_time_ninja.get_attribute("value") + "_" + module + ".png"))
    
    url_signout = "https://www.ninjacrm.com/"+client_name+"/index.php?logout=true"
    driver.get(url_signout)

    



for i in range(len(df)):
    login(i)
    ss(i)

