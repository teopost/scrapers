#!/usr/local/bin/python3

# pip3 install bs4
# pip3 install lxml
# pip3 install xlsxwriter
# pip3 install selenium
# pip3 install pandas

# brew install geckodriver

# https://hackernoon.com/building-a-web-scraper-from-start-to-finish-bb6b95388184

from bs4 import BeautifulSoup
import requests

from lxml.html import fromstring
from itertools import cycle
import traceback

import xlsxwriter 
import time
import os
import random

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import pandas as pd

val = {
    "Nome": 0,
    "Cognome": 1,
    "PEC": 2,
    "Settori": 3,
    "Data di nascita": 4,
    "Citta' di nascita": 5,
    "Indirizzo dello Studio": 6,
    "Data Iscrizione": 7,
    "Laurea Citta'": 8,
    "Laurea Data": 9,
}

'''
    "località di abilitazione": 10,
    "anno di abilitazione": 11,
    "località prima iscrizione albo": 12,
    "data prima iscrizione albo": 13,
    "Commissario per esami": 14,
    "Nominato presso amministrazioni, enti pubblici, società, ecc. per": 15,
    "numero di matricola": 16,
    "partita IVA": 17,
    "indirizzo dello studio": 18,
    "telefono dello studio": 19,
    "Diploma di maturità": 20,
    "Universitari e post universitari": 21,
    "Modalità di svolgimento della professione di architetto": 22,
    "Progettazione architettonica": 23,
    "Progettazione di Interni": 24,
    "cellulare": 25,
    "Attività diversa": 26,
    "Settore di appartenenza del dipendente pubblico": 27,
    "denominazione studio": 28,
    "presso": 29,
    "fax dello studio": 30,
    "Libero professionista (solo per attività inerenti la professione di architetto)": 31,
    "Attività specialistiche": 32,
'''

user_agent_list = [
   #Chrome
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
    'Mozilla/5.0 (Windows NT 5.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.2; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.90 Safari/537.36',
    'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.157 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/57.0.2987.133 Safari/537.36',
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
    #Firefox
    'Mozilla/4.0 (compatible; MSIE 9.0; Windows NT 6.1)',
    'Mozilla/5.0 (Windows NT 6.1; WOW64; Trident/7.0; rv:11.0) like Gecko',
    'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)',
    'Mozilla/5.0 (Windows NT 6.1; Trident/7.0; rv:11.0) like Gecko',
    'Mozilla/5.0 (Windows NT 6.2; WOW64; Trident/7.0; rv:11.0) like Gecko',
    'Mozilla/5.0 (Windows NT 10.0; WOW64; Trident/7.0; rv:11.0) like Gecko',
    'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0)',
    'Mozilla/5.0 (Windows NT 6.3; WOW64; Trident/7.0; rv:11.0) like Gecko',
    'Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; Trident/5.0)',
    'Mozilla/5.0 (Windows NT 6.1; Win64; x64; Trident/7.0; rv:11.0) like Gecko',
    'Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; WOW64; Trident/6.0)',
    'Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)',
    'Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 5.1; Trident/4.0; .NET CLR 2.0.50727; .NET CLR 3.0.4506.2152; .NET CLR 3.5.30729)'
]

#opts = webdriver.FirefoxOptions()
#opts.headless = True
#driver = webdriver.Firefox(options=opts)
driver = webdriver.Firefox()
driver.implicitly_wait(30)


def get_person_page(url):

    user_agent = random.choice(user_agent_list)
    headers = {'User-Agent': user_agent}
    response = requests.get(url, timeout=25, headers=headers)

    soup = BeautifulSoup(response.content, "lxml")
    rows = soup.find("div", attrs={"id": "itemListSecondary"})

    link_collection=[]
    for link in rows.find_all("a"):
        link_collection.append('https://www.archiworld-fc.it' + link.get("href"))
        
    return link_collection

def get_person_details(url):

    #user_agent = random.choice(user_agent_list)
    #headers = {'User-Agent': user_agent}
    #response = requests.get(url, headers=headers, timeout=25)
    #print (response.text)

    #soup = BeautifulSoup(response.content, "lxml")
      
    driver.get(url)

    soup=BeautifulSoup(driver.page_source, 'lxml')

    my_list = ()

    status = True
    try:
        # estraggo il nome
        l1 = soup.find("ul", attrs={"class": "uk-list"})
        for detail in l1.find_all("li"):
            label = detail.text.split(":")[0].strip()

            if label == "PEC":
                value1 = detail.text.split(":")[1].strip()
                value =  value1.split()[0]
            else:
                value = detail.text.split(":")[1].strip()

            my_list += (  [ label , value  ], )

    except Exception as e:
            print (e)
            status = False

    return my_list, status

def get_col(key):
    try:
        index = val[key]
    except KeyError as e:
        index = 99

    return index


if __name__ == "__main__":

    #rowdata, stato = get_person_details('https://www.ordinearchitetti.mi.it/it/ordine/albo/scheda/-72997043-roberto-avanzini')

    workbook = xlsxwriter.Workbook('scraping.xlsx') 
    worksheet = workbook.add_worksheet("My sheet")
    bold = workbook.add_format({'bold': True})

    row=0
    # stampo intestazione
    for intestazione in (val): 
        col=get_col(intestazione)
        worksheet.write(row, col, intestazione.upper(), bold)
        col += 1 

    row=1

    # Puoi arrivare fino a pagina 810 (27 pagine)
    beginpage=0
    endpage=817
    

    try:
        # 840
        for i in range(beginpage, endpage, 30):
            currentpage=i

            mainurl = 'https://www.archiworld-fc.it/index.php?option=com_k2&view=itemlist&task=filter&searchword74=A%20-%20a%20architetto&moduleId=245&Itemid=423&limitstart=' + str(i)
            print ('>> Get page: ' + mainurl)

            links = get_person_page(mainurl)
            
            for page in links:
                print ('  - Parsing: ' + page)
                #time.sleep(5)
                rowdata, stato = get_person_details(page)
                if stato == True:
                    # stampo valori
                    col=0
                    for intestazione, valore in (rowdata): 
                        col=get_col(intestazione)
                        if col != 99:
                            worksheet.write(row, col, valore)
                            col += 1 
                    row += 1
                else:
                    print ("ERROR: PARSING FAILED!!!: " + page)
            time.sleep(5)
    except Exception as e:
        print (e)
        pass
    finally:
        workbook.close()
    
    os.rename('scraping.xlsx', 'scraping-' + str(beginpage) + '-' + str(endpage) + '.xlsx')
    driver.close()
    print ("DONE")
    



