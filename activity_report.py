#!/usr/bin/python

# Author: Clara Watson
# Date: August, 2022
# Description:  this program is designed for National Tax Advisory Services 
# to more efficiently pull data from each of their 15 databases this program
# pulls the activity reports as csv files and stores them to the local computer
 
from selenium import webdriver 
import time
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import os
import glob
import shutil


def setup():
    """
    this function stores the names and links to all the databases 
    so it can loop through them 
    """
    nat = "https://nattax.irslogics.com"
    tag = "https://taxag.irslogics.com"
    ntas = "https://ntas.irslogics.com"
    taa = "https://taxaa.irslogics.com"
    grnlw = "https://greenlawtax.irslogics.com"
    tj = "https://taxoffice.irslogics.com"
    boca = "https://brtax.irslogics.com"
    hnst = "https://honest.logiqs.com"
    gtrl = "https://gtrl.irslogics.com"
    ss = "https://stateside.irslogics.com"
    honestNV = "https://honestnv.irslogics.com"
    nte = "https://nte.irslogics.com"
    ovation = "https://ovationtax.irslogics.com"
    hurricane = "https://hurricanetax.logiqs.com"
    recovery = "https://recoverytax.logiqs.com"
    logics = [nat, tag, ntas, taa, grnlw, tj, boca, hnst, gtrl, ss, honestNV, nte, ovation, hurricane, recovery]
    name = ["nat", "tag", "ntas", "taa", "grnlw", "tj", "boca", "hnst", "gtrl", "ss", "honestNV", "nte", "ovation", "hurricane", "recovery"]
    browser = webdriver.Chrome(r"C:\Users\Clara\AppData\Local\Temp\Temp1_chromedriver_win32.zip\chromedriver.exe")
    time.sleep(3) 
    # loops through the logics array that stores the links 
    for i in range(0,len(logics)): 
        # opens each link in browser
        browser.get(logics[i])
        time.sleep(5)
        # calls the next function
        login(browser, name[i],logics[i] )
    
def login(browser, name, url): 
    """
    this function is called after the setup function opens 
    a database login page and sends the users login information
    to the webpage
    """
    email = ""
    pswd = ""
    # fills in password and username by sending keys to the browser
    username = browser.find_element(By.ID,'txtUsername2')
    username.clear()
    username.send_keys(email)
    password = browser.find_element(By.ID,'txtPassword2')
    password.clear()
    password.send_keys(pswd)
    # clicks the login button
    browser.find_element(By.ID, 'btnLogin2').click()
    time.sleep(5) 
    # wait until the "get reports" button is visible after loading and clicks it 
    WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div[1]/div/div[1]/div[2]/div[2]/ul/li[6]"))).click()
    # goes directly to the activity reports
    browser.get(url + "/Reports/ActivityReport.aspx")
    time.sleep(5)
    # calls next function to download the report
    get_report(browser, name)
    
def get_report(browser, name ):
    """
    this function fills in the information needed to access 
    the activity report and executes the download script to 
    save the report as a csv file 
    """
    start = "08/23/2022"
    time.sleep(2)
    # fills in the report start and end date by sending keys 
    start_date= browser.find_element(By.ID, "txtStartDate")
    start_date.clear()
    start_date.send_keys(start)
    end_date = browser.find_element(By.ID, "txtEndDate")
    end_date.clear()
    end_date.send_keys(start)
    time.sleep(3)
    # clicks on the "generate report" button
    browser.find_element(By.ID, 'btnGenerate').click() 
    time.sleep(3)
    # calls a built in javascript function to export the file
    browser.execute_script("exportexcel()")
    time.sleep(4)
    print("downloaded")
    # calls the next function to edit file 
    downloaded(browser, start, name)
    
def downloaded(browser, start,name ):
    """
    this function finds the most recently downlaoded file
    and renames it according to what database the report is 
    from and saves it to a new directory 
    """
    file_name = "REPORT_" + name + "_" + start.replace("/", "-") + ".csv"
    print(file_name)
    browser.find_element(By.ID, 'btnGenerate').click() 
    # using os and glob to get a list of recently downloaded files
    home = os.path.expanduser("~")
    downloadspath = os.path.join(home, "Downloads")
    list_of_files = glob.glob(downloadspath + "\*.csv") 
    # gets the most recently downloaded file aka the report 
    latest_file = max(list_of_files, key=os.path.getctime)
    print(latest_file)
    # copies the file to a new folder and renames it 
    shutil.copy(str(latest_file), r"C:\Users\Clara\Documents\Reports_Local" + "\\" + file_name )
    print("copied and renamed file")
    return

    
if __name__ == '__main__':
    setup()
    
