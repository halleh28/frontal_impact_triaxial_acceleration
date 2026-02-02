"""
Created on Mon Jul 22 10:11:03 2019

@author: Matthew Sie (original)
Edited 12/5/2019 by Caitlin McCleery
Edited 11/4/2025 by Halle Holzbauer
"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as ec
from selenium.common.exceptions import TimeoutException
import time

# Using Chrome to access web, set window to fullscreen, specify path to location of chromedriver file
chromeOptions = Options()
chromeOptions.add_argument("--start-maximized")
driver = webdriver.Chrome(options=chromeOptions)

# Open the IIHS website
driver.get('https://techdata.iihs.org')

wait = WebDriverWait(driver, 15)

wait.until(ec.element_to_be_clickable((By.LINK_TEXT, "Log in to IIHS TechData"))).click()

wait.until(ec.presence_of_element_located((By.ID, "cred_userid_inputtext"))).send_keys("hholzbauer@nationalbi.com")
driver.find_element(By.ID, "cred_password_inputtext").send_keys("1013266XIa!p?")
driver.find_element(By.ID, "cred_sign_in_button").click()

wait.until(ec.element_to_be_clickable((By.LINK_TEXT, "Continue to downloads"))).click()

dropdown = Select(wait.until(ec.presence_of_element_located(
    (By.NAME, "ctl00$MainContentPlaceholder$TestTypeDropdown")
)))
dropdown.select_by_visible_text("Frontal - updated test")
driver.find_element(By.ID, "ctl00_MainContentPlaceholder_SearchButton").click()

time.sleep(2)

#links = driver.find_elements_by_xpath("//a[@href]")
cases = []
page_dropdown_name = "ctl00$MainContentPlaceholder$PageSelector$ddPages"
#for link in links:
#        if 'filegroup' in link.get_attribute("href"):
#            cases.append([link.get_attribute("href"), link.get_attribute('text')])
#            print ('Adding Case: ' + link.get_attribute('text'))
   
for x in range(1,9):  #range of page tabs to go through - range(n,m) goes n to m-1 (not m)
    Select(driver.find_element(By.NAME, page_dropdown_name)).select_by_visible_text(str(x))
    time.sleep(2)

    links = driver.find_elements(By.XPATH, "//a[@href]")
    for link in links:
        href = link.get_attribute("href")
        text = link.text.strip()
        if "filegroup" in href:
            cases.append([href, text])
            print(f"Adding Case: {text}")

print(f"\nTotal cases found: {len(cases)}")

#    cases = [[link.get_attribute("href"), link.get_attribute('text')] for link in driver.find_elements_by_xpath("//a[@href]") if 'filegroup' in link.get_attribute("href")]
#    [print ('Adding Case: ' + link.get_attribute('text')) for link in driver.find_elements_by_xpath("//a[@href]") if 'filegroup' in link.get_attribute("href")]

for idx, case in enumerate(cases):
    case_name = case[1][:7]
    print("\n************************************")
    print(f"DOWNLOADING REPORTS FOR: {case_name}")
    print(case)
    
    driver.get(case[0])

    wait.until(ec.element_to_be_clickable((
        By.XPATH, "//*[contains(translate(text(), 'abcdefghijklmnopqrstuvwxyz', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'), 'DIADEM')]"
    ))).click()

    wait.until(ec.presence_of_all_elements_located((By.TAG_NAME, "a")))
    time.sleep(2)

    tdms_links = driver.find_elements(
    By.XPATH,
    "//a[contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '.tdms') "
    "and not(contains(translate(text(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '.tdms_index'))]"
)

    if not tdms_links:
        print(f"No .tdms files found for {case_name}")
    else:
        print(f"Found {len(tdms_links)} .tdms files for {case_name}")
        for i, elem in enumerate(tdms_links, start=1):
            try:
                wait.until(ec.element_to_be_clickable(elem))
                try:
                    elem.click()
                except:
                    driver.execute_script("arguments[0].click();", elem)
                print(f"→ Downloaded .tdms file {i}/{len(tdms_links)} for {case_name}")
                time.sleep(2)
            except Exception as e:
                print(f"Could not download .tdms file {i}: {e}")

    print(f"Cases downloaded: {idx + 1}/{len(cases)}")
    print("************************************\n")

print("\nAll available Frontal - original test cases processed successfully.")
#    if data_excel1:
#        try:    
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'plots')))
#        finally:
#            driver.find_element_by_link_text('plots').click()
#    elif data_excel2:
#        try:    
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'Summary')))
#        finally:
#            driver.find_element_by_link_text('Summary').click()
    
#    try:
#        element=driver.find_element_by_partial_link_text('plots')
#    except NoSuchElementException:
#        print('text not found')
              
    
 #   data_excel1 = driver.find_element_by_partial_link_text('plots').click()
#    data_excel2 = driver.find_element_by_partial_link_text('Summary')
    
    #data_excel1 = driver.find_elements_by_xpath("//*[contains(text(), 'plots')]")
    #data_excel2 = driver.find_elements_by_xpath("//*[contains(text(), 'Summary')]")
    
#    if data_excel1:
#        try:    
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'plots')))
#        finally:
#            driver.find_element_by_link_text('plots').click()
#    elif data_excel2:
#        try:    
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'Summary')))
#        finally:
#            driver.find_element_by_link_text('Summary').click()
    
#    data_excel1 = driver.find_elements_by_xpath("//*[contains(text(), 'DATA\EXCEL')]")
#    data_excel2 = driver.find_elements_by_xpath("//*[contains(text(), 'Data\EXCEL')]")
#    
#    if data_excel_upper:
#        try:    
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'DATA\EXCEL')))
#        finally:
#            driver.find_element_by_link_text('DATA\EXCEL').click()
#    elif data_excel_lower:
#        try:    
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'Data\EXCEL')))
#        finally:
#            driver.find_element_by_link_text('Data\EXCEL').click()
#        
#    if data_excel_lower or data_excel_upper:
#        summary = True
#        
#        try:
#            file = WebDriverWait(driver, 45).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'Summary')))
#        except TimeoutException:
#            summary = False
#            pass
#        
#        if not summary:
#            try:
#                file = WebDriverWait(driver, 45).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, 'plots')))
#            except TimeoutException:
#                pass
#    
#        if not summary:
#            file = driver.find_element_by_partial_link_text('plots')
#        else:
#            file = driver.find_element_by_partial_link_text('Summary')
#        print ('Downloading plots')
#        actionChain = ActionChains(driver)
#        actionChain.context_click(file).perform()

#    time.sleep(0.5)
#    pyautogui.press('down')
#    time.sleep(0.5)
#    pyautogui.press('down')
#    time.sleep(0.5)
#    pyautogui.press('down')
#    time.sleep(0.5)
#    pyautogui.press('down')
#    time.sleep(1)
#    pyautogui.press('enter')
#    
#    time.sleep(1.5)
#    pyautogui.press('tab')
#    pyautogui.press('tab')
#    time.sleep(0.5)
#    pyautogui.press('tab')
#    pyautogui.press('tab')
#    pyautogui.press('tab')
#    pyautogui.press('tab')
#    time.sleep(0.5)
#    if count == 0:
#        pyautogui.press('right')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        pyautogui.press('enter')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        pyautogui.press('down')
#        pyautogui.press('down')
#        pyautogui.press('down')
#        pyautogui.press('down')
#        pyautogui.press('down')
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        pyautogui.press('down')
#        pyautogui.press('down')         
#        pyautogui.press('down')
#        pyautogui.press('enter')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('enter')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('enter')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('enter')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('up')
#        time.sleep(0.5)
#        pyautogui.press('enter')        
#    else:
#        pyautogui.press('right')
#        pyautogui.press('right')
#        time.sleep(0.5)
#        pyautogui.press('right')
#        pyautogui.press('right')
#        time.sleep(0.5)
#        pyautogui.press('right')
#        pyautogui.press('right')
#        pyautogui.press('right')
#        time.sleep(0.5)
#        pyautogui.press('enter')        
#
#    time.sleep(0.5)
#    pyautogui.hotkey('ctrl','shift','n')
#    time.sleep(0.5)
#    pyautogui.typewrite(case_name)
#    time.sleep(0.5)
#    pyautogui.press('enter')
#    time.sleep(0.5)
#    pyautogui.press('enter')
#    time.sleep(0.5)
#    pyautogui.press('tab')  
#    pyautogui.press('tab')   
#    time.sleep(0.5)
#    pyautogui.press('tab')   
#    pyautogui.press('tab')   
#    time.sleep(0.5)
#    pyautogui.press('tab')   
#        
#    time.sleep(0.5)
#    pyautogui.press('enter')             
#
#    driver.switch_to.window(parent_h)
#
#    # Download Report
#
#    driver.get(case[0])
#    time.sleep(1)
#    
#    has_report = driver.find_elements_by_xpath("//*[contains(text(), 'REPORTS')]")
#    
#    if has_report:
#        try:
#            excel = WebDriverWait(driver, 120).until(ec.visibility_of_element_located((By.LINK_TEXT, 'REPORTS')))
#        finally:
#            driver.find_element_by_link_text('REPORTS').click()
#        try:
#            report = WebDriverWait(driver, 115).until(ec.visibility_of_element_located((By.PARTIAL_LINK_TEXT, '.pdf')))
#        finally:
#            report = driver.find_element_by_partial_link_text('.pdf')
#            print ('Downloading report')
#            actionChain = ActionChains(driver)
#            actionChain.context_click(report).perform()
#        
#        time.sleep(0.75)
#        pyautogui.press('down')
#        time.sleep(0.75)
#        pyautogui.press('down')
#        time.sleep(0.75)
#        pyautogui.press('down')
#        time.sleep(0.5)
#        pyautogui.press('down')
#        time.sleep(1)
#        pyautogui.press('enter')
#        time.sleep(1)
#        pyautogui.press('enter')
#        time.sleep(0.5)
#          
    