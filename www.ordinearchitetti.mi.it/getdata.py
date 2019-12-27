#!/usr/local/bin/python3

# pip3 install bs4
# pip3 install lxml
# pip3 install xlsxwriter

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

val = {
    "titolo": 0,
    "nome": 1,
    "email": 2,
    "sito": 3,
    "nato/a a": 4,
    "il": 5,
    "cod. fisc.": 6,
    "data iscrizione": 7,
    "località di laurea": 8,
    "anno di laurea": 9,
    "località di abilitazione": 10,
    "anno di abilitazione": 11,
    "località prima iscrizione albo": 12,
    "data prima iscrizione albo": 13,
}

'''
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

def get_proxies():
    url = 'https://free-proxy-list.net/'
    response = requests.get(url)
    parser = fromstring(response.text)
    proxies = set()
    for i in parser.xpath('//tbody/tr')[:10]:
        if i.xpath('.//td[7][contains(text(),"yes")]'):
            #Grabbing IP and corresponding PORT
            proxy = ":".join([i.xpath('.//td[1]/text()')[0], i.xpath('.//td[2]/text()')[0]])
            proxies.add(proxy)
    return proxies

def get_person_page(page, prx):
    # 1210
    url = 'https://www.ordinearchitetti.mi.it/it/ordine/albo/' + str(page)

    #Pick a random user agent
    user_agent = random.choice(user_agent_list)
    #Set the headers 
    headers = {'User-Agent': user_agent}

    if prx == "":
        response = requests.get(url, timeout=25, headers=headers)
    else:
        response = requests.get(url, timeout=25, headers=headers, proxies={"http": prx, "https": prx})

    soup = BeautifulSoup(response.content, "lxml")

    #rows = soup.find("div", attrs={"class": "archialbo"})

    rows = soup.find("div", attrs={"id": "wraparchialbo"})

    link_collection=[]
    
    for link in rows.find_all("a"):
        #print("title: {}".format(link.get("title")))
        if not link.get("title"):
            
            link_collection.append('https://www.ordinearchitetti.mi.it' + link.get("href"))
            #print("href: https://www.ordinearchitetti.mi.it{}".format(link.get("href")))

    return link_collection

def get_person_details(url, prx):

    #Pick a random user agent
    user_agent = random.choice(user_agent_list)
    #Set the headers 
    headers = {'User-Agent': user_agent}

    if prx == "":
        response = requests.get(url, headers=headers, timeout=25)
    else:
        response = requests.get(url, headers=headers, timeout=25, proxies={"http": prx, "https": prx})
            
    soup = BeautifulSoup(response.content, "lxml")

    my_list = ()

    dict = { } 
    status = True
    try:
        # estraggo il nome
        l1 = soup.find("div", attrs={"id": "datip"})
        for detail in l1.findAll('h1', attrs={"class": "h3"}):
            #junk = detail.text.encode('utf-8').strip()
            junk = detail.text.strip()
            dict['titolo'] = junk.splitlines()[0].strip()
            dict['nominativo'] = junk.splitlines()[1].strip()

            my_list += (  ['titolo' , junk.splitlines()[0].strip()], )
            my_list += (  ['nome' , junk.splitlines()[1].strip()], )

        # estraggo i campi

        rows = soup.find("div", attrs={"class": "datiPersonali"})
        if rows != None:
            status = True
            for detail in rows.find_all("p"):      
                for span in detail.find_all("span"):
                    if span.nextSibling != None:
                        my_list += ( [span.text.strip(), span.nextSibling.strip() ], )
                    
                        dict[span.text.strip()] = span.nextSibling.strip()

                    #print (span.text.encode('utf-8').strip())
                    #print (span.nextSibling.encode('utf-8').strip())
    except:
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


    # PROXY IMPLEMENTATION
    #proxies = get_proxies()
    #  NOOOOO 
    #proxies = {
    #    '209.90.63.108:80'
    #}

    #proxies = {}
    #proxy_pool = cycle(proxies)
    proxy = ""

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
    # Puoi arrivare fino a pagina 1210
    beginpage=1

    offset=10
    currentpage=0
    try:
        for i in range(beginpage, beginpage+offset):
            #proxy = next(proxy_pool)
            currentpage=i
            print ('>> Get page: ' + str(i) + ' of ' + str(beginpage+offset))
            links = get_person_page(i, proxy)
            
            for page in links:
                print ('  - Parsing: ' + page)
                #time.sleep(5)
                rowdata, stato = get_person_details(page, proxy)
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
    
    os.rename('scraping.xlsx', 'scraping-' + str(beginpage) + '-' + str(currentpage) + '.xlsx')
    print ("DONE")
    



