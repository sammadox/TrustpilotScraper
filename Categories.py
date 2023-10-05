from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import datetime
from datetime import date
from openpyxl import Workbook
import sys
from selenium.webdriver.common.keys import Keys
import os
from selenium.webdriver.common.action_chains import ActionChains



#Save to workbook
def save_to_workbook(L):
    wb = Workbook()

    sheet = wb.active

    sheet['A1']="Sport"
    sheet['J1']="Link"
    sheet['B1']="Country"
    sheet['C1']="URL"
    sheet['E1']="Pick"
    sheet['G1']="Combo_Pick"
    sheet['H1']="FORMULA"
    sheet['I1']="timestamp"
    sheet['K1']="closing odds"
    sheet['L1']="Outcome"
    sheet['F1']="event_date"
    sheet['M1']="Stake"
    sheet['N1']="Odds"
    sheet['D1']="event_name"
    
    H=['A','B','C','D','E','F','G','H','I','J','K','L','M','N']
    idx=2
    for o in L:
        try:
            for i in range(len(o)):
                cell=str(H[i])+str(idx)
                sheet[cell]=o[i]
            idx+=1
        except:
            print('error has been occured')

    # Save the workbook
    wb.save(filename='Results.xlsx')

#Split the values from the main menu 
def split_string_by_newline(input_string):
    # Split the input string into a list of substrings using '\n' as the separator
    substrings = input_string.split('\n')
    return substrings

#To initiate the driver
def init_driver():
    chrome_driver_path = "chromedriver.exe"


    options = webdriver.ChromeOptions()
    options.binary_location = "ungoogled-chromium_114.0.5735.110-1.1_windows/chrome.exe"  # Replace with the path to the Chromium binary if needed
    
    # Initialize Chrome driver instance
    driver = webdriver.Chrome(service=ChromeService(executable_path=chrome_driver_path),options=options)

    #return driver
    return(driver)

#To expand the list of matches
def expand_match(page_link,driver):
    
    # Navigate to the url
    driver.get(page_link)

    #Get all cards 
    scroll_pixels = 500
    driver.execute_script(f'window.scrollBy(0, {scroll_pixels});')
    L=driver.find_elements(By.XPATH,"/html/body/div[1]/div/div/main/div/section/div[3]/div/div[*]")
    
    
    ex_data(driver,1,5)
    while(True):
        #Close the driver
        #driver.quit()
        #Close the code
        #sys.exit()
        pass

#To check how much page I do have
def n_page(driver):
    #Get the number of results
    #Once opened click on statistics by category
    if int(driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div[2]/p").text.split(" ")[2])%20==0:
        n_res=int(driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div[2]/p").text.split(" ")[2])//20
    else:
        n_res=int(driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div[2]/p").text.split(" ")[2])//20+1
  
    time.sleep(2)
    scrolls(driver,1)
    print(n_res)

#A functon to visit a given page
def nav_page(driver,pg):
    lk="https://www.trustpilot.com/categories/animal_health?page="+str(pg)
    driver.get(lk)
#For each page extract relevant data
def ex_data(driver,nb,ls_nb):
    if nb==ls_nb:
        pass
    else:
        #Make it less zoomed
        #Click on the e-mail section
        #Accept cookies usage
        try:
            driver.find_element(By.XPATH,"/html/body/div[2]/div[2]/div/div[1]/div/div[2]/div/button[2]").click()
        except:
            try:
                time.sleep(1)
                driver.find_element(By.XPATH,"/html/body/div[2]/div[2]/div/div[1]/div/div[2]/div/button[2]").click()
            except:
                pass
        L=[]
        for o in range(4,24):
            H=[]
            
            driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/div[2]").click()
            #Now we're going to extract all data
            #Get the business name
            b_n=""
            try:
               
                b_n=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div[2]/p").text
            except:
                try:
                    b_n=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div/div[2]/p").text
                except:
                    pass
            H.append(b_n)
            #Get the business image src
            b_i=""
            try:
                b_i=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div[1]/picture/img").get_attribute("src")
            except:
                try:
                    b_i=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div/div[1]/picture/img").get_attribute("src")
                except:
                    pass
            H.append(b_i)
            #Get the trustscore
            t_s=""
            try:
                t_s=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div[2]/div[1]/p/span[1]").text
            except:
                try:
                    t_s=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div/div[1]/p/span[1]").text
                except:
                    pass
            H.append(t_s)
            #Number of reviews
            n_b=""
            try:
                n_b=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div[2]/div[1]/p").text
            except:
                try:
                    n_b=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div/div[1]/p").text
                except:
                    pass
            H.append(n_b)
            #Country
            c=""
            try:
                c=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div[2]/div[2]/span").text
            except:
                try:
                    c=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/a/div[2]/div/div[2]/span").text
                except:
                    pass
            H.append(c)
            #Extract the List of keywords
            kwds=""
            try:
                kwds=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/div[2]/div/div").text
            except:
                try:
                    kwds=driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/div/div/div").text
                except:
                    pass
            H.append(kwds)
            driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/div[2]/div/span/button").click()
            #Extract the website
            web=""
            try:
                web=driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[2]/div[2]/span/ul/li[1]").text
            except:
                pass
            H.append(web)
            #Extract e-mail
            em=""
            try:
                em=driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[2]/div[2]/span/ul/li[2]").text
            except:
                pass
            H.append(em)
            #Extract Phone
            ph=""
            try:
                ph=driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[2]/div[2]/span/ul/li[3]").text
            except:
                pass
            H.append(ph)
            #Extrcat address
            add=""
            try:
                add=driver.find_element(By.XPATH,"/html/body/div[4]/div/div/div[2]/div[2]/span/ul/li[4]").text
            except:
                pass
            H.append(add)
            driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/div[2]").click()

            print(H)
            driver.find_element(By.XPATH,"/html/body/div[1]/div/div/main/div/div[2]/div/section/div["+str(o)+"]/div[2]").click()
            L.append(H)
        print(L)

           
#Create a function that crolls
def scrolls(driver, seg):
    if seg==1:
        scroll_pixels = 500
    driver.execute_script(f'window.scrollBy(0, {scroll_pixels});')
    
driver=init_driver()
#Call the function
expand_match('https://www.trustpilot.com/categories/animal_health',driver)
