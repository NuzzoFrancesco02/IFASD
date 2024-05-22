import pdfplumber
import numpy as np
import re
from openpyxl import Workbook, load_workbook
import os 
from scholarly import scholarly
from scholarly import ProxyGenerator
import socks
import socket
import random
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from pdfminer.high_level import extract_text

proxies = ['http://38.242.158.32:443',
    'http://185.105.91.159:4444',
    'http://185.217.199.114:4444',
    'http://195.235.124.143:80',
    'http://45.9.75.240:4444',
    'http://84.252.75.63:4444',
    'http://185.128.106.91:4444',
    'http://5.75.206.99:80',
    'http://5.45.110.13:80',
    'http://185.128.106.40:4444',
    'http://43.128.112.143:3128',
    'http://185.128.106.80:4444',
    'http://185.128.107.49:4444',
    'http://46.232.248.164:80',
    'http://61.158.175.38:9002',
]
def random_delay():
    time.sleep(random.uniform(1, 5))

def get_random_proxy():
    """Returns a random proxy from the list."""
    return random.choice(proxies)

def cerca(book_title):
    """Searches for a book on Google Scholar and returns the link to the first result."""
    # Setup Chrome options for headless browsing
    proxy = get_random_proxy()
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    #options.add_argument("--proxy-server=socks5://127.0.0.1:9050")
    options.add_argument(
        "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
    )

    # Initialize the WebDriver
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        # Open Google Scholar
        driver.get("https://scholar.google.com/")
        random_delay()

        # Locate the search box and enter the book title
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(book_title)
        random_delay()
        search_box.send_keys(Keys.RETURN)

        random_delay()

        # Get the search results
        results = driver.find_elements(By.CLASS_NAME, "gs_ri")
        #print(f"Found {len(results)} results")  # Debug: print number of results found

        # Check if there are any results
        if not results:
            print("No results found")
            return None

        # Extract the link from the first result
        for idx, result in enumerate(results, 1):
            try:
                title_element = result.find_element(By.TAG_NAME, "h3")
                link = title_element.find_element(By.TAG_NAME, 'a').get_attribute('href')
                snippet = result.find_element(By.CLASS_NAME, 'gs_rs').text

                # Return the link of the first result
                if idx > 1:
                    return link
            except NoSuchElementException as e:
                print(f"Error in result {idx}: {e}")

    except NoSuchElementException as e:
        print(f"Error during search: {e}")
        return None

    finally:
        # Ensure the driver quits properly
        driver.quit()
def clean_text(paragrafo):
    if '\n' in paragrafo:
        paragrafo = re.sub(r'(?<=[a-zA-Z])\n|(?=\n[a-zA-Z])',' ',paragrafo)
        paragrafo = re.sub(r'\n+','',paragrafo)
    if '  ' in paragrafo:
        pat = r'\s+'
        paragrafo = re.sub(pat,' ',paragrafo)
    if '  ' in paragrafo:
        pat = r'\s+'
        paragrafo = re.sub(pat,' ',paragrafo)
    if '- ' in paragrafo:
        paragrafo = re.sub('- ','',paragrafo)
    return(paragrafo)
#print(cerca('Cross-modal Damping Model: an experimental extraction approach and aircraft dynamic loads application'))
wb = load_workbook('FINAL PAPER IFASD 2017/Repository Upload Information_IFASD_2017.xlsx')

ws = wb.active
for row in np.arange(9,171,1):
    flag_abstract = flag_keywords = flag_online = True
    if ws['G'+str(row)].value == None:
        flag_abstract = False
    if ws['I'+str(row)].value == None:
        flag_keywords = False
    if ws['J'+str(row)].value == None:
        flag_online = False
    
    titles = []
    
    if not flag_online or not flag_keywords or not flag_abstract:
        is_already_open = False
        IFASD_name = ws['F'+str(row)].value
        #pdf_path = os.path.join('FINAL PAPER IFASD 2017',IFASD_name)
        texts = extract_text('FINAL PAPER IFASD 2017/'+IFASD_name,page_numbers=[0,1])
        #extract_text(

        #with pdfplumber.open(pdf_path) as pdf:
        if not flag_abstract:
            if not is_already_open:

                #first_page = pdf.pages[0]
                
                #second_page = pdf.pages[1]
                #pattern = r'\bK\s*e\s*y\s*w\s*o\s*r\s*d\s*s\b'
                '''
                if not re.search(pattern, paragrafo, re.IGNORECASE):
                    first_page = pdf.pages[1]
                    paragrafo = first_page.extract_text_simple(x_tolerance=1, y_tolerance=5)
                '''
                is_already_open = True
            #paragrafo1 = first_page.extract_text_simple(x_tolerance=1, y_tolerance=5)
            #paragrafo2 = second_page.extract_text_simple(x_tolerance=1, y_tolerance=5)
            #paragrafo = paragrafo1 + ' ' + paragrafo2
            paragrafo = texts
            pattern = r'Abstract:\s*(.*?)[\s*\n]?\s*INTRODUCTION'
            #pattern = r'Abstract:\s*(.*?)\s*1\s*INTRODUCTION'
            paragrafo = clean_text(paragrafo)
            try:
                match = re.search(pattern, paragrafo,re.DOTALL)
                paragrafo = match.group(1)  
            except:
                try: 
                    pattern = r'Abstract:\s*(.*?)[\s*\n]?\s*List of Symbols'
                    match = re.search(pattern, paragrafo,re.DOTALL)
                    paragrafo = match.group(1)         
                except:
                    try:
                        pattern = r'Abstract:\s*(.*?)[\s*\n]?\s*Introduction'
                        match = re.search(pattern, paragrafo,re.DOTALL)
                        paragrafo = match.group(1) 
                    except:       
                        pattern = r'Abstract:\s*(.*?)[\s*\n]?\s*Notice to Readers'
                        match = re.search(pattern, paragrafo,re.DOTALL)
                        paragrafo = match.group(1) 
            pat = r'1\s+'+IFASD_name
            if pat in paragrafo:
                paragrafo = re.sub(pat,'',paragrafo)
            
            #print(paragrafo + '\n\n')
            ws['G'+str(row)].value = paragrafo
        
        if not flag_keywords and IFASD_name!='IFASD-2017-180.pdf' and IFASD_name!='IFASD-2017-181.pdf':
            if not is_already_open:
                #first_page = pdf.pages[0]                   
                #second_page = pdf.pages[1]
                is_already_open = True
            #paragrafo1 = first_page.extract_text_simple(x_tolerance=1, y_tolerance=1)
            #paragrafo2 = second_page.extract_text_simple(x_tolerance=1, y_tolerance=1)
            #paragrafo = paragrafo1 + ' ' + paragrafo2
            paragrafo = texts
            pattern =  r'Keywords:\s*(.*?)\s*[\n\s]*Abstract\s*:'
            #paragrafo = first_page.extract_text_simple(x_tolerance=5, y_tolerance=5)
            paragrafo = clean_text(paragrafo)
            try:
                match = re.search(pattern, paragrafo,re.DOTALL)   
                paragrafo = match.group(1)
            except:
                try:
                    pattern =  r'Kewywords:\s*(.*?)[\s*\n]?\s*INTRODUCTION'
                    match = re.search(pattern, paragrafo,re.DOTALL)   
                    paragrafo = match.group(1)
                except:
                    try:
                        pattern =  r'Kewywords:\s*(.*?)[\s*\n]?\s*Introduction'
                        match = re.search(pattern, paragrafo,re.DOTALL)   
                        paragrafo = match.group(1)
                    except:
                        pattern =  r'Keywords:\s*(.*?)\s*1\s+INTRODUCTION'
                        match = re.search(pattern, paragrafo,re.DOTALL)   
                        paragrafo = match.group(1)
            
            
            #print(paragrafo + '\n\n')
            ws['G'+str(row)].value = paragrafo
        if not flag_online or flag_online:
            
            title = ws['B'+str(row)].value
            if '\n' in title:          
                title = re.sub(r'\n+',' ',title)
                
            if '  ' in title:
                pat = r'\s+'
                title = re.sub(pat,' ',title)
            
            #title = 'Comparison of High-Fidelity Aero-Structure Gradient Computation Techniques. Application on the CRM Wing Design.'
            print('\n'+title)
            #title = re.sub(' ','+',title)
            #site = "https://scholar.google.com/scholar?hl=it&as_sdt=0%2C5&q="+title+".&btnG="
            #link = cerca(title)
            #if link == None:
            #    link = 'checked'
            #print(link)
            print('\n')
            titles = titles.append(title+'\n')
            #ws['J'+str(row)].value = link
with open('titles.txt', "w") as file_titles:
    file_titles.write(titles)             
wb.save('FINAL PAPER IFASD 2017/Repository Upload Information_IFASD_2017.xlsx')