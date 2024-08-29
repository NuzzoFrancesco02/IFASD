import numpy as np
import re
from openpyxl import Workbook, load_workbook
import os 
import random
import time
from pdfminer.high_level import extract_text



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
    pat = r'1\s+' + re.escape(IFASD_name)
    if re.search(pat, paragrafo):
        paragrafo = re.sub(pat,'',paragrafo)
    return(paragrafo)
wb = load_workbook('/Users/francesconuzzo/Downloads/USB_Proceedings/info/papers/Cartel1.xlsx')

ws = wb.active
titles = ''
raw_begin = 147
raw_end = 210

if isinstance(ws['E'+str(raw_begin - 1)].value, int):
    indice = ws['E'+str(raw_begin - 1)].value + 1
else:
    indice = raw_begin-8
for row in np.arange(raw_begin,raw_end,1):
    if indice == 209:
        break
    flag_authors = flag_conf_numb = flag_abstract = flag_file_name = flag_keywords = flag_category = flag_online = True
    flag_title = False
    #if ws['B'+str(row)].value == None:
        #flag_title = False
    if ws['C'+str(row)].value == None:
        flag_authors = False
    if ws['E'+str(row)].value == None:
        flag_conf_numb = False
    if ws['F'+str(row)].value == None:
        flag_file_name = False
    if ws['G'+str(row)].value == None:
        flag_abstract = False
    if ws['H'+str(row)].value == None:
        flag_category = False
    if ws['I'+str(row)].value == None:
        flag_keywords = False
    if ws['J'+str(row)].value == None:
        flag_online = False
    
    

    
    if not flag_title or not flag_authors or not flag_conf_numb or not flag_file_name or not flag_category or not flag_online or not flag_keywords or not flag_abstract:
        is_already_open = False
        is_file_flag = False
        while not is_file_flag and indice <= 208:
            if indice < 10:
                IFASD_name = 'IFASD-2015-00'+str(indice)+'.pdf'
                if os.path.isfile('/Users/francesconuzzo/Downloads/PROCEEDINGS/'+IFASD_name):
                    
                    is_file_flag = True
                else:
                    indice = indice + 1
            elif indice < 100:
                IFASD_name = 'IFASD-2015-0' + str(indice)+'.pdf'
                if os.path.isfile('/Users/francesconuzzo/Downloads/PROCEEDINGS/'+IFASD_name):
                
                    is_file_flag = True
                else:
                    indice = indice + 1
            else:
                IFASD_name = 'IFASD-2015-' + str(indice)+'.pdf'
                if os.path.isfile('/Users/francesconuzzo/Downloads/PROCEEDINGS/'+IFASD_name):
                    
                    is_file_flag = True
                else:
                    indice = indice + 1
        if not flag_conf_numb:
                ws['E'+str(row)].value = indice
        if is_file_flag == True:
            indice = indice + 1
        #IFASD_name = ws['F'+str(row)].value
        if not 'IFASD-2019-063.pdf' == IFASD_name and not 'IFASD-2019-069.pdf' == IFASD_name and not 'IFASD-2019-100.pdf' == IFASD_name and not 'IFASD-2019-129.pdf' == IFASD_name:
        #if not 'IFASD-2015-054.pdf' == IFASD_name and not 'IFASD-2015-181.pdf' == IFASD_name and not 'IFASD-2015-206.pdf' == IFASD_name and not 'IFASD-2015-207.pdf' == IFASD_name:
            if not flag_title:
                if not is_already_open:
                    texts = extract_text('/Users/francesconuzzo/Downloads/USB_Proceedings/info/papers/'+IFASD_name,page_numbers=[0,1])
                    is_already_open = True
                    paragrafo = texts
                    
                    #paragrafo = clean_text(paragrafo,IFASD_name)
                #uppercase check
                up_title = paragrafo[:10]
                up_title_check = clean_text(up_title,IFASD_name)
                up_title_check = up_title_check.replace(" ","")
                
                if up_title_check.isupper():
                    #pattern = r"(^[^a-z]*)([a-z].*)"
                    pattern =r'\s*(.*?)(?:\r?\n\r?\n)'
                    match = re.search(pattern, paragrafo,re.DOTALL)
                    paper_title = match.group(1)
                    paper_title = paper_title[:-2]
                    beg, title_pos = match.regs[1]
                else:
                    #pattern =r'(.*?)(?:\r?\n\r?\n)'
                    pattern =r'\s*(.*?)(?:\r?\n\r?\n)'
                    match = re.findall(pattern,paragrafo,re.DOTALL)
                    paper_title = match[0]
                    title_pos = len(paper_title)

                ws['B'+str(row)].value = clean_text(paper_title,IFASD_name)
            if not flag_authors:
                if not is_already_open:
                    texts = extract_text('/Users/francesconuzzo/Downloads/USB_Proceedings/info/papers/'+IFASD_name,page_numbers=[0,1])
                    is_already_open = True
                    paragrafo = texts
                    #paragrafo = clean_text(paragrafo,IFASD_name)
                #pattern = re.escape(paper_title) + r'\n*([^\n]+)'
                pattern = re.escape(paper_title) + r'\n*(.+?)(?=\n{2,})'
                match = re.findall(pattern, paragrafo, re.DOTALL)

                if match:
                    paper_authors = match[0]
                #paper_authors = re.sub(pattern,'', paper_authors)
                #authors_pos = title_pos + len(paper_authors) 
                paper_authors = re.sub(r'\d+', '', paper_authors)
                ws['C'+str(row)].value = clean_text(paper_authors,IFASD_name)

            
            if not flag_file_name:
                ws['F'+str(row)].value = IFASD_name
            if not flag_abstract:
                if not is_already_open:
                    texts = extract_text('/Users/francesconuzzo/Downloads/USB_Proceedings/info/papers/'+IFASD_name,page_numbers=[0,1])
                    is_already_open = True
                paragrafo = texts
                paragrafo = clean_text(paragrafo,IFASD_name)
                pattern = r'Abstract:\s*(.*?)[\s*\n]?[\d]\s*INTRODUCTION'
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
                            try:
                                pattern = r'Abstract:\s*(.*?)[\s*\n]?\s*Notice to Readers'
                                match = re.search(pattern, paragrafo,re.DOTALL)
                                paragrafo = match.group(1) 
                            except:
                                try:
                                    pattern = r'Abstract:?\s*(.*?)\s*\n*\s*INTRODUCTION'
                                    match = re.search(pattern, paragrafo,re.DOTALL)
                                    paragrafo = match.group(1) 
                                except:
                                    try:
                                        pattern = r'Abstract:?\s*(.*?)\s*\n*\s*INTRODUCYION'
                                        match = re.search(pattern, paragrafo,re.DOTALL)
                                        paragrafo = match.group(1) 
                                    except:
                                        try:
                                            pattern = r'Abstract:?\s*(.*?)\s*\n*\s*NOMENCLATURE'
                                            match = re.search(pattern, paragrafo,re.DOTALL)
                                            paragrafo = match.group(1) 
                                        except:
                                            try:
                                                pattern = r'Abstract:?\s*(.*?)\s*\n*\s*NOTATION'
                                                match = re.search(pattern, paragrafo,re.DOTALL)
                                                paragrafo = match.group(1) 
                                            except:
                                                try: 
                                                    pattern = r'Abstract.\s*(.*?)[\s*\n]?[\d]\s*INTRODUCTION'
                                                    match = re.search(pattern, paragrafo,re.DOTALL)
                                                    paragrafo = match.group(1) 
                                                except:
                                                    try: 
                                                        pattern = r'Abstract:?\s*(.*?)\s*\n*\s*Nomenclature'
                                                        match = re.search(pattern, paragrafo,re.DOTALL)
                                                        paragrafo = match.group(1) 
                                                    except: 
                                                        try: 
                                                            pattern = r'Abstract:?\s*(.*?)\s*\n*\s*1 BACKGROUND AND MOTIVATION'
                                                            match = re.search(pattern, paragrafo,re.DOTALL)
                                                            paragrafo = match.group(1) 
                                                        except:
                                                            try:
                                                                pattern = r'Abstract:?\s*(.*?)\s*\n*\s*1 LOADING OF CONTROL SURFACE ACTUATORS'
                                                                match = re.search(pattern, paragrafo,re.DOTALL)
                                                                paragrafo = match.group(1) 
                                                            except:
                                                                pattern = r'Abstract:?\s*(.*?)\s*\n*\s*1 INFLUENCE OF FLIGHT'
                                                                match = re.search(pattern, paragrafo,re.DOTALL)
                                                                paragrafo = match.group(1) 
                                                                
                    
                    

                                                        

                    if paragrafo[-1].isdigit():
                        paragrafo = paragrafo[:-1]
                ws['G'+str(row)].value = paragrafo     
            
            #if not flag_keywords and IFASD_name!='IFASD-2017-180.pdf' and IFASD_name!='IFASD-2017-181.pdf':
                if not is_already_open:
                    IFASD_name = ws['F'+str(row)].value
                    texts = extract_text('/Users/francesconuzzo/Downloads/USB_Proceedings/info/papers/'+IFASD_name,page_numbers=[0,1])
                    is_already_open = True
                paragrafo = texts
                pattern =  r'Keywords:\s*(.*?)\s*[\n\s]*Abstract\s*:'
                paragrafo = clean_text(paragrafo,IFASD_name)
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
                            try:
                                pattern =  r'Keywords:\s*(.*?)\s*[\n\s]*Abstract '
                                match = re.search(pattern, paragrafo,re.DOTALL)   
                                paragrafo = match.group(1)
                            except:
                                try:
                                    pattern =  r'Keywords:\s*(.*?)\s*1\s+INTRODUCTION'
                                    match = re.search(pattern, paragrafo,re.DOTALL)   
                                    paragrafo = match.group(1)
                                except:
                                    try:
                                        pattern =  r'Keywords.\s*(.*?)[\s*\n]?\s*Abstract'
                                        match = re.search(pattern, paragrafo,re.DOTALL)   
                                        paragrafo = match.group(1)
                                    except:
                                        try:
                                            pattern =  r'Keywords\s*(.*?)[\s*\n]?\s*Abstract'
                                            match = re.search(pattern, paragrafo,re.DOTALL)   
                                            paragrafo = match.group(1)
                                        except:
                                            try: 
                                                pattern =  r'Key words:\s*(.*?)[\s*\n]?\s*Abstract'
                                                match = re.search(pattern, paragrafo,re.DOTALL)   
                                                paragrafo = match.group(1)
                                            except:
                                                pattern =  r'Kewords:\s*(.*?)[\s*\n]?\s*Abstract'
                                                match = re.search(pattern, paragrafo,re.DOTALL)   
                                                paragrafo = match.group(1)
                    
                    if paragrafo[-1].isdigit():
                        paragrafo = paragrafo[:-1]
                ws['I'+str(row)].value = paragrafo
            if not flag_online or flag_online:
                
                title = ws['B'+str(row)].value
                if '\n' in title:          
                    title = re.sub(r'\n+',' ',title)
                    
                if '  ' in title:
                    pat = r'\s+'
                    title = re.sub(pat,' ',title)
                print('\n'+title)
                print('\n')
                print(IFASD_name + '\n')

                titles = titles + title +'\n'
                #ws['J'+str(row)].value = link
        else:
            row = row + 1
with open('titles.txt', "w") as file_titles:
    file_titles.write(titles)             
wb.save('/Users/francesconuzzo/Downloads/USB_Proceedings/info/papers/Cartel1.xlsx')