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

#Calculate 6 Months from now
def calculate_six_months_from_now():
    # Get the current date
    current_date = datetime.date.today()
    # Calculate six months from now
    six_months_from_now = current_date - datetime.timedelta(days=30*3)
    return six_months_from_now
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
#Scrape Data from each page
def scrape_fm_p(p_link,ori_link):

    chrome_driver_path = "chromedriver.exe"

    options = webdriver.ChromeOptions()
    options.binary_location = "ungoogled-chromium_114.0.5735.110-1.1_windows/chrome.exe"  # Replace with the path to the Chromium binary if needed
        
    # Initialize Chrome driver instance
    driver = webdriver.Chrome(service=ChromeService(executable_path=chrome_driver_path),options=options)



    print("1")
    #Initialize the empty list
    L=[]
    L.append("soccer")
    #Go to the link
    driver.get(p_link)

    #Get the country#Correct
    try:
        ctry=driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[4]/div[2]/span[3]/a").text.split(', ')[1]
    except:
        ctry='-'
    L.append(ctry)
 
    #Get the URL#Correct
    L.append(ori_link)

    #Extract Event Name
    try:
        ev_name=driver.find_element(By.XPATH,'/html/body/div[1]/div[4]/div[4]/div[2]/h1').text
    except:
        ev_name='-'
    L.append(ev_name)

    #Get the Pick#Correct
    try:
        pick=driver.find_element(By.XPATH,'/html/body/div[1]/div[4]/div[4]/div[2]/span[2]').text.split(': ')[1].replace("Cuota","")
    except:
        pick='-'
    L.append("'"+split_string_by_newline(pick)[0])

    #Extract the event date#Correct
    try:
        evdt=driver.find_element(By.XPATH,'/html/body/div[1]/div[4]/div[4]/div[2]/span[3]').text
    except:
        evdt='-'
    td=split_string_by_newline(evdt)

    

    #Extract date#Correct
    try:
        ev_dt=td[1].split(": ")[1]
    except:
        ev_dt='-'
    L.append(ev_dt)


    #Extract combo_pick#correct
    co_p='TBC'
    L.append(co_p)

    #Extract Formula#Correct
    fo_v='TBC'
    L.append(fo_v)

    #Extract Timestamp#Correct
    ti_v='TBC'
    L.append(ti_v)

    #Get the link#Correct
    lk=driver.current_url
    L.append(lk)

    #Extract closing Odds#Correct
    clo_v='TBC'
    L.append(clo_v)

    

    #Extract outcome#Correct
    try:
        ou_v=ev_name.split('(')[1].replace(" ", "").replace("c", "").replace(")", "")
    except:
        ou_v='-'
    L.append(ou_v)

    
    print("2")

    #Extract stake
    try:
        st_dt=td[3].split(": ")[1]
    except:
        st_dt='-'
    L.append(st_dt)

    

    #Extract odds
    try:
        od_v=split_string_by_newline(driver.find_element(By.XPATH,'/html/body/div[1]/div[4]/div[4]/div[2]/span[2]').text)[1].split(': ')[1]
    except:
        od_v='-'
    L.append(od_v)



















   
    
    

    

    
    

    

    

    

    

    
    

   
    

   

    print(L)
    return(L)
#Format the date
def format_date(input_date):
    year=int(input_date.split('-')[2])
    month=int(input_date.split('-')[1])
    day=int(input_date.split('-')[0])
    fr=date(year,month,day)
    return(fr)
#Calculate six months before
def calculate_six_months_before_today(today_date):
    # Convert the input string to a datetime object
    today_datetime = datetime.strptime(today_date, '%d-%m-%Y')
    # Calculate the date 6 months before today
    six_months_ago = today_datetime.replace(day=1) - timedelta(days=30)
    # Format the result as DD-MM-YYYY
    formatted_date = six_months_ago.strftime('%d-%m-%Y')
    
    return formatted_date
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

    #Once opened click on statistics by category
    element=driver.find_element(By.XPATH,"/html/body/div[1]/div[4]/div[4]/div[2]/select")

    time.sleep(2)
    #Count all + occurences
    L=driver.find_elements(By.XPATH,"/html/body/div[1]/div[4]/div[4]/div[2]/div[4]/table/tbody/tr[*]/td[6]/a")
    for i in L:
        try:
            i.click()
        except:
            print("An exception occurred on expand match")
    
    #Open all hrefs under the pick class
    L=op_hrefs(driver)
    print(L)
    RES=[]
    for i in range(0,len(L),2):
        try:
            RES.append(scrape_fm_p(L[i],page_link))
        except:
            pass
    save_to_workbook(RES)
    

    while(True):
        #Close the driver
        driver.quit()
        #Close the code
        sys.exit()
        pass
    
    # Close the driver
    driver.quit()
#Open all href inside the class picks
def op_hrefs(driver):
    P=[]
    D=[]
    L=driver.find_elements(By.TAG_NAME,"tr")
    dt=''
    for i in L:
        if '[+]' in i.text:
            print(i.text)
            yr=i.text.split('-')[1]
        H=i.find_elements(By.TAG_NAME,"img")
        for o in H:
            if o.get_attribute("src")=="https://www.tipgol.com/sports/templates/images/sports/soccer.png":
                MP=i.find_elements(By.TAG_NAME,"a")
                idx=0
                date_lim=calculate_six_months_from_now().strftime("%d-%m-%Y")

                date_lim=format_date(date_lim)
                for mp in MP:
                    dt=split_string_by_newline(i.text)[idx]+'-'+yr.split()[0]
                    dt=format_date(dt)
                    if mp.get_attribute("href") not in P and 'php' not in mp.get_attribute("href") and dt>date_lim :
                        P.append(mp.get_attribute("href"))
                        P.append(split_string_by_newline(i.text)[idx]+'-'+yr.split()[0])
                        idx+=1
                        break
                        print(P)
                        if len(P)==10:
                            break
                    else:
                        if dt<date_lim:
                            break
                if dt<date_lim:
                    break
            if dt!='':
                if dt<date_lim:
                    break
        if dt!='':
            if dt<date_lim:
                break
    return(P)
driver=init_driver()
#Call the function
expand_match('https://www.tipgol.com/tipster,1517.html',driver)
