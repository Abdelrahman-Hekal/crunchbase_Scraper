from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait as wait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import csv
import os
from datetime import datetime
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException
import pandas as pd
import undetected_chromedriver.v2 as uc
import warnings
warnings.filterwarnings('ignore')
import unidecode
import pickle
import requests
from requests_ip_rotator import ApiGateway, EXTRA_REGIONS
import random

def initialize_bot():

    # Setting up chrome driver for the bot
    #chrome_options  = webdriver.ChromeOptions()
    chrome_options = uc.ChromeOptions()
    #chrome_options.add_argument('--headless')
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
    chrome_options.add_argument('--log-level=3')
    chrome_options.add_argument("--enable-javascript")
    chrome_options.add_argument("--start-maximized")
    chrome_options.add_argument("--disable-blink-features")
    chrome_options.add_argument("--incognito")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument('--disable-gpu')
    ########################################
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-impl-side-painting")
    chrome_options.add_argument("--disable-setuid-sandbox")
    chrome_options.add_argument("--disable-seccomp-filter-sandbox")
    chrome_options.add_argument("--disable-breakpad")
    chrome_options.add_argument("--disable-client-side-phishing-detection")
    chrome_options.add_argument("--disable-cast")
    chrome_options.add_argument("--disable-cast-streaming-hw-encoding")
    chrome_options.add_argument("--disable-cloud-import")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--disable-session-crashed-bubble")
    chrome_options.add_argument("--disable-ipv6")
    chrome_options.add_argument("--allow-http-screen-capture")
    chrome_options.add_argument("--start-maximized")

    chrome_options.add_argument("--disable-extensions") 
    chrome_options.add_argument("--disable-notifications") 
    chrome_options.add_argument("--disable-infobars") 
    chrome_options.add_argument("--remote-debugging-port=9222")
    chrome_options.add_argument('--disable-dev-shm-usaging')
############################################
    #path = ChromeDriverManager().install()
    #driver = webdriver.Chrome(path, options=chrome_options)
    chrome_options.page_load_strategy = 'normal'
    driver = uc.Chrome(options=chrome_options)

    return driver

def login(driver, name, pwd):
    URL1 = "https://www.crunchbase.com/home"
    # navigating to the website link
    driver.get(URL1)
    time.sleep(5)
    URL1 = "https://www.crunchbase.com/login"
    # navigating to the website link
    driver.get(URL1)
    time.sleep(5)
    # signing in 
    username = driver.find_element_by_name("email")
    password = driver.find_element_by_name("password")
    username.send_keys(name)
    time.sleep(1)
    password.send_keys(pwd)
    time.sleep(1)
    driver.find_element_by_xpath("//button[@type='submit' and @class='mat-focus-indicator login mat-raised-button mat-button-base mat-primary']").click()
    time.sleep(5)

def scrape_data(driver, name):

    company = '{} Alumni Founded Companies'.format(name)
    #company = 'LinkedIn Alumni Founded Companies'
    driver.get('https://www.crunchbase.com/discover/hubs')
    time.sleep(5)
    search = driver.find_element_by_xpath("//input[@type='search' and @id='mat-input-0']")
    search.click()
    search.send_keys(company)
    time.sleep(5)
    #elements = driver.find_elements_by_tag_name('search-results-section')
    #time.sleep(2)
    #element = elements[2]
    try:
        #driver.find_element_by_xpath('/html/body/chrome/div/mat-sidenav-container/mat-sidenav-content/div/multi-search-results/page-layout/div/div/div/div/search-results-section[3]/mat-card/a').click()
        wait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "/html/body/chrome/div/mat-sidenav-container/mat-sidenav-content/div/multi-search-results/page-layout/div/div/div/div/search-results-section[3]/mat-card/a"))).click()
    except:
        return False, 'NA'
    time.sleep(2)
    company_site = driver.current_url
    #search_res = driver.find_element_by_xpath('/html/body/chrome/div/mat-sidenav-container/mat-sidenav-content/div/ng-component/entity-v2/page-layout/div/div/profile-header/div/header/div/div/div/div[2]/div[1]/h1').text
    search_res = wait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "/html/body/chrome/div/mat-sidenav-container/mat-sidenav-content/div/ng-component/entity-v2/page-layout/div/div/profile-header/div/header/div/div/div/div[2]/div[1]/h1"))).text

    res = search_res.replace(' Alumni Founded Companies', '').lower().strip()
    input = name.lower().strip()
    #if res != input:
    if res not in input and input not in res:
        return False, res

    parent = driver.find_element_by_xpath("//div[@class='identifier-label']").text  
    path = 'D:\Hekal\Personal\Freelancing\Scraping_companies_info\scraped_data\\' + parent
    if not os.path.isdir(path):
        os.makedirs(path)

    sections = driver.find_elements_by_xpath("//row-card[@class='ng-star-inserted']")
    time.sleep(2)
    card_labels = ['leaderboard', 'investors', 'funding', 'investments', 'acquisitions', 'people']
    links = []
    header = []
    row = []
    for i, section in enumerate(sections):
        # Overview section
        if i == 0:
            parent = driver.find_element_by_xpath("//div[@class='identifier-label']").text    
            header.append('Parent')
            row.append(parent)
            norg = driver.find_element_by_xpath("/html/body/chrome/div/mat-sidenav-container/mat-sidenav-content/div/ng-component/entity-v2/page-layout/div/div/div/page-centered-layout[2]/div/div/div[1]/row-card[1]/profile-section/section-card/mat-card/div[2]/div/fields-card[1]/ul/li[2]/field-formatter/a").text
            header.append('Number of Organizations')
            row.append(norg)
            #print(norg)
            ul = section.find_elements_by_xpath("//ul[@class='text_and_value']")[1]
            li = ul.find_elements_by_tag_name('li')
            time.sleep(2)
            for i in range(len(li)):
                if 'Number of Founders' in li[i].text:
                    nfounders = li[i].text.split('\n')[1].replace(',', '')
                    header.append('Number of Founders')
                    row.append(nfounders)
                    #print('Num of founders')
                    #print(nfounders)
                if 'Average Founded Date' in li[i].text:
                    founded_date = li[i].text.split('\n')[1]  
                    header.append('Average Founded Date')
                    row.append(founded_date)                    
                    #print('founded date')
                    #print(founded_date)
                if 'Percentage Acquired' in li[i].text:
                    per_acquired = li[i].text.split('\n')[1]   
                    header.append('Percentage Acquired')
                    row.append(per_acquired) 
                    #print('Percentage Acquired')
                    #print(per_acquired)
                if 'Percentage of Public Organizations' in li[i].text:
                    per_public = li[i].text.split('\n')[1]    
                    header.append('Percentage of Public Organizations')
                    row.append(per_public)
                    #print('Percentage of Public Organizations')
                    #print(per_public)
                if 'Percentage Non-Profit' in li[i].text:
                    per_nonprofit = li[i].text.split('\n')[1] 
                    header.append('Percentage Non-Profit')
                    row.append(per_nonprofit)
                    #print('Percentage Non-Profit')
                    #print(per_nonprofit)                        
                if 'Top Investor Types' in li[i].text:
                    top_investors = li[i].text.split('\n')[1]
                    header.append('Top Investor Types')
                    row.append(top_investors)
                    #print('Top Investor Types')
                    #print(top_investors)                
                if 'Number of For-Profit Companies' in li[i].text:
                    ncomp = li[i].text.split('\n')[1]
                    header.append('Number of For-Profit Companies')
                    row.append(ncomp)
                    #print('Number of For-Profit Companies')
                    #print(ncomp)

        # IPO section
        elif section.find_element_by_tag_name('h2').text.lower().strip() == 'ipo':
            nIPOs = section.find_elements_by_tag_name("a")[0].text
            header.append('Number of IPOs')
            row.append(nIPOs)
            #print('Number of IPOs')
            #print(nIPOs)
            total_IPO = section.find_elements_by_tag_name("a")[1].text
            header.append('Total Amount Raised in IPO')
            row.append(total_IPO)
            #print('Total Amount Raised in IPO')
            #print(total_IPO)
            ul = section.find_elements_by_tag_name("ul")[0]
            li = ul.find_elements_by_tag_name('li')
            time.sleep(2)
            for i in range(len(li)):
                if 'Median Amount Raised in IPO' in li[i].text:
                    IPO_amount = li[i].text.split('\n')[1].replace(',', '')
                    header.append('Median Amount Raised in IPO')
                    row.append(IPO_amount)
                    #print('Median Amount Raised in IPO')
                    #print(IPO_amount)
                if 'Total IPO Valuation' in li[i].text:
                    total_IPO_valuation = li[i].text.split('\n')[1]   
                    header.append('Total IPO Valuation')
                    row.append(total_IPO_valuation)
                    #print('Total IPO Valuation')
                    #print(total_IPO_valuation)
                if 'Median IPO Valuation' in li[i].text:
                    median_IPO_valuation = li[i].text.split('\n')[1]   
                    header.append('Median IPO Valuation')
                    row.append(median_IPO_valuation)
                    #print('Median IPO Valuation')
                    #print(median_IPO_valuation)
                if 'Average IPO Date' in li[i].text:
                    avg_IPO_date = li[i].text.split('\n')[1]    
                    header.append('Average IPO Date')
                    row.append(avg_IPO_date)
                    #print('Average IPO Date')
                    #print(avg_IPO_date)
                if 'Percentage Delisted' in li[i].text:
                    per_delisted = li[i].text.split('\n')[1] 
                    header.append('Percentage Delisted')
                    row.append(per_delisted)
                    #print('Percentage Delisted')
                    #print(per_delisted)   
 
                    
        # Leaderboard, Investor, Funding, Investment, Acquisitions and People sections
        else: 
           label = section.find_element_by_tag_name('h2').text.lower().strip()
           for name in card_labels:
               if name in label:
                   #elems = driver.find_elements_by_xpath("//a[@class='mat-focus-indicator mat-button mat-button-base mat-primary ng-star-inserted']")
                   #link = elems[i - 1].get_attribute('href')
                   elems = section.find_elements_by_tag_name('a')
                   link_found = False
                   for elem in elems:
                       if elem.text.lower() == 'view all':
                           link = elem.get_attribute('href')
                           links.append(link)
                           link_found = True
                           break

                   if not link_found:
                       list_card = section.find_element_by_tag_name('list-card')
                       table = list_card.find_element_by_tag_name('table')
                       html = table.get_attribute('outerHTML')
                       scrape_mini_tabular_data(html, parent, path, label.capitalize())


                   if 'funding' in label:
                       try:
                           card = section.find_elements_by_tag_name('big-values-card')[0]
                           row.append(card.find_elements_by_tag_name('a')[0].text)
                           header.append(card.find_elements_by_tag_name('label-with-info')[0].text)
                       except:
                            missing = True

                   elif 'investors' in label:
                       try:
                           card = section.find_elements_by_tag_name('big-values-card')[0]
                           row.append(card.find_elements_by_tag_name('a')[0].text)
                           header.append(card.find_elements_by_tag_name('label-with-info')[0].text)
                       except:
                            missing = True

                   elif 'investments' in label:
                       try:
                           card = section.find_elements_by_tag_name('big-values-card')[0]
                           row.append(card.find_elements_by_tag_name('a')[0].text)
                           header.append(card.find_elements_by_tag_name('label-with-info')[0].text)
                       except:
                            missing = True

                   elif 'acquisitions' in label:
                       try:
                           card = section.find_elements_by_tag_name('big-values-card')[0]
                           row.append(card.find_elements_by_tag_name('a')[0].text)
                           header.append(card.find_elements_by_tag_name('label-with-info')[0].text)
                       except:
                            missing = True
                            
                   elif 'People' in label:
                       try:
                           card = section.find_elements_by_tag_name('big-values-card')[0]
                           row.append(card.find_elements_by_tag_name('a')[0].text)
                           header.append(card.find_elements_by_tag_name('label-with-info')[0].text)
                       except:
                            missing = True


    with open (path + '\\' + parent + '_Overview_IPO.csv', 'w', newline='') as file:
        csvwriter = csv.writer(file)
        csvwriter.writerow(header)
        csvwriter.writerow(row)

    for i, link in enumerate(links):
        scrape_tabular_data(driver, link, parent, path)

    return True, res
    #time.sleep(100)

def scrape_tabular_data(driver, link, company, path):
    driver.get(link)
    time.sleep(5)
    table_name = driver.find_elements_by_tag_name('h2')[0].text
    table_name = table_name.split(' ')[0]
    print(table_name)
    print('-'*50)

    # getting table rows data
    parent = driver.find_element_by_xpath("//section-card[@class='ng-star-inserted']")
    child = parent.find_element_by_xpath("//list-card[@class='full-width ng-star-inserted']") 
    while True:
        try:
            button = child.find_elements_by_xpath("//span[@class='mat-button-wrapper']")[-2]
            if 'show more' not in button.text.lower():
                break
            driver.execute_script("arguments[0].click();", button)
            time.sleep(2)
        except:
            break
    table = driver.find_element_by_xpath("//table[@class='card-grid']")
    html = table.get_attribute('outerHTML')
    html = unidecode.unidecode(html)
    df = pd.read_html(html)[0]
    df['Parent'] = company
    df.to_csv(path + '\\' + company + '_' + table_name +'.csv', index=False, encoding='UTF-8')
    
def scrape_mini_tabular_data(table, company, path, name):
    table_name = name.split(' ')[0]
    print(table_name)
    print('-'*50)
    html = unidecode.unidecode(table)
    df = pd.read_html(html)[0]
    df['Parent'] = company
    df.to_csv(path + '\\' + company + '_' + table_name +'.csv', index=False, encoding='UTF-8')


def process_data():

    data = {}
    cwd = os.getcwd()
    path = cwd + '\\scraped_data'
    folders = os.listdir(path) 
    for folder in folders:
        files_path = path + '\\' + folder
        files =  os.listdir(files_path) 
        for file in files:
            df = pd.read_csv(files_path+ '\\' + file, encoding='UTF-8')
            filename = file.split('_')[-1]
            filename = filename.lower().capitalize()
            if data.get(filename, 0) == 0:
                data[filename] = [df.copy()]
            else:
                data[filename].append(df.copy())

    output_path = cwd + '\\Final output\\'
    if not os.path.isdir(output_path):
        os.mkdir(output_path)

    for key, values in data.items(): 
        df = pd.DataFrame()
        for value in values:
            if df.shape[0] > 0:
                df = df.append(value, ignore_index = True)
            else:
                df = value
        
        cols = df.columns
        ord_cols = []
        ord_cols.append(cols[-1])
        for i in range (len(cols)-1):
            ord_cols.append(cols[i])
        df = df[ord_cols]
        if 'Funding' in key:
            df['Series'] = df['Transaction Name'].apply(lambda x: x.split('-')[0].replace('Series', '').replace('Round', ''))
            df['Organization Name'] = df['Transaction Name'].apply(lambda x: x.split('-')[-1])
            df.drop('Transaction Name', axis=1, inplace=True)
            cols = df.columns
            ord_cols[:2] = cols[:2]
            ord_cols[2:4] = cols[-2:]
            ord_cols[4:6] = cols[2:4]
            df = df[ord_cols]        
        elif 'Investments' in key:
            df['Series'] = df['Funding Round'].apply(lambda x: x.split('-')[0].replace('Series', '').replace('Round', ''))
            df['Organization Name'] = df['Funding Round'].apply(lambda x: x.split('-')[-1])
            df.drop('Funding Round', axis=1, inplace=True)
            cols = df.columns
            ord_cols[:3] = cols[:3]
            ord_cols[3:5] = cols[-2:]
            ord_cols[5:6] = cols[3:4]
            df = df[ord_cols]        
        elif 'Acquisitions' in key:
            df['Organization Name'] = df['Transaction Name'].apply(lambda x: x.split('acquired')[0])
            df.drop('Transaction Name', axis=1, inplace=True)
            cols = df.columns
            ord_cols[:3] = cols[:3]
            ord_cols[3:4] = cols[-1:]
            ord_cols[4:5] = cols[3:4]
            df = df[ord_cols]        
        elif 'ipo' in key.lower():
            cols = df.columns
            ord_cols[:] = cols[1:]
            ord_cols.append(cols[0])
            df = df[ord_cols]
            key = 'Overview_IPO.csv'
        #########################################################
        # for mapping gvkey data
        df2 = pd.read_csv('status_found.csv')
        nrows = df.shape[0]
        df['gvkey'] = 0
        for i in range(nrows):
            comp = df.iloc[i, 0]
            gvkey = int(df2[df2['Result'] == comp]['gvkey'].values[0])
            df.loc[i, 'gvkey'] = gvkey

        cols = df.columns
        ord_cols[1:] = cols[:-1]
        ord_cols[0] = cols[-1]
        df = df[ord_cols]

        df.to_csv(output_path + 'Companies_' + key, sep=',', encoding='utf-8', index=False)

def clear_screen():
    # for windows
    if os.name == 'nt':
        _ = os.system('cls')
  
    # for mac and linux
    else:
        _ = os.system('clear')

    
# Website login credentials
name1 = ""
pwd1 = ""


os.system("color")
COLOR = {"HEADER": "\033[95m", "BLUE": "\033[94m", "GREEN": "\033[92m",
            "RED": "\033[91m", "ENDC": "\033[0m", "YELLOW" :'\033[33m'}


if __name__ == '__main__':
    if os.path.exists('fail.txt'):
        os.remove('fail.txt')
    driver = initialize_bot()
    clear_screen()
    nfail = 0
    signin = True
    exclude = ['inc', 'corp', 'co', 'plc', 'cp', 'ltd', 'nv', 'sa', 'llc', 'lp', '.com']
    try:
        df = pd.read_excel('companies.xlsx')
    except:
        pass
    df2 = pd.read_excel('companies_full.xlsx')
    comp = df.iloc[:, 1].values.tolist()
    ncomp = len(comp)
    df_scraped = pd.read_csv('status.csv')
    nscraped = df_scraped.shape[0]
    for i in range(nscraped, ncomp):
        code = df2[df2['conm'] == comp[i]]['gvkey'].values[0]
        name_ = comp[i].split('-')[0]
        words = name_.split(' ')
        n = len(words)
        for j in range(n):
            if words[j].lower() in exclude:
                words[j] = ''

        name = ' '.join(words)
        name = name.replace('.com', '').replace('.COM', '')
        print(name)
        try:
            if signin:
                login(driver, name1, pwd1)
                signin = False
            found, res = scrape_data(driver, name)
            with open ('status.csv', 'a', newline='') as file:
                csvwriter = csv.writer(file)
                csvwriter.writerow([code, comp[i], res, found])
            row = "The data has been fetched successfully for company number {}".format(i+1)
            print(COLOR["GREEN"],row, COLOR["ENDC"])
            time.sleep(5)
            nfail = 0
            time.sleep(5)
            with open('fail.txt', 'w') as file:
                write = True
        except:
            row ='failure in accessing the website!'
            print(COLOR["RED"],row, COLOR["ENDC"])
            exit()

    process_data()
