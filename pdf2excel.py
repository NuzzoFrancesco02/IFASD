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
import requests
from bs4 import BeautifulSoup


def cerca_link_google(title):
    title = re.sub(' ','+',title)
    url = 'https://scholar.google.com/scholar?hl=it&as_sdt=0%2C5&q='+title+'&btnG='
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"}
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")
        links = soup.find_all("a", href=True)
        for link in links:
            href = link.get("href")
            print(href)

cerca_link_google('A method for predicting multivariate random loads and a discrete appoximation of the mutidimensional design load envelope')


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
def cerca(book_title):
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get("https://scholar.google.com/")
        random_delay()

        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys(book_title)
        random_delay()
        search_box.send_keys(Keys.RETURN)

        random_delay()
        results = driver.find_elements(By.CLASS_NAME, "gs_ri")

        for idx, result in enumerate(results, 1):
            title_element = result.find_element(By.TAG_NAME, "h3")
            title = title_element.text
            link = title_element.find_element(By.TAG_NAME, 'a').get_attribute('href')
            snippet = result.find_element(By.CLASS_NAME, 'gs_rs').text

            if idx > 1:
                break
            
            #print(f"Link: {link}")
            return link
           
    except NoSuchElementException:
        return 'checked'
        
    finally:
        driver.quit()

def clean_text(paragrafo,IFASD_name):
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
    IFASD_name=re.sub('.pdf','',IFASD_name)
    pat = r'1\s+' + re.esca