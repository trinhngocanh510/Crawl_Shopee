# Import webdriver
from http import client
from operator import index
from turtle import title
import pygsheets
import gspread
from bs4 import BeautifulSoup
from argparse import Action
from lib2to3.pgen2 import driver
from telnetlib import EC
from xml.dom.minidom import Element
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from oauth2client.service_account import ServiceAccountCredentials
from soupsieve import select
from selenium.webdriver.edge.service import Service
import time
import re


#Initial sheet
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
cred = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json',scope)
client = gspread.authorize(cred)
sheet = client.open("Shopee").sheet1

initial_title = ['Running...........','................','................','Loading...........','................','................']
title = ['Product name','Type of shop','Quantity sold','Location','Discount','Freeship','Shop rating']

# Running
sheet.update('A1:F1',[initial_title])
sheet.format('A1:F1', {'textFormat':{'italic':True}})


# Initial driver

driverPath = '.\msedgedriver.exe'
s = Service(driverPath)
driver = webdriver.Edge(service=s)
driver.maximize_window()
url = "https://shopee.vn/search?keyword="
    
listProduct = open('nameProduct.txt').readlines()


# Initial value
products_details = []
initial_number = 1
number_page_scan = 1

# Write to sheet
def write_to_sheet(product_details):
    global initial_number
    initial_number +=1
    sheet.update(f'A{initial_number}:G{initial_number}',[product_details])
    sheet.format(f'A{initial_number}:G{initial_number}', {'textFormat':{'italic':True}})
    time.sleep(1.2)
    
    
    
# Scan web
def scan_web(page_source):   
    for product in page_source:
        
        # Get name product
        products_name = product.find('div', class_ = 'ie3A+n bM+7UW Cve6sh').get_text().strip()
        
        # Get type shop
        type_shop = ''
        if 'YeGYFd sKFCYs' in str(product):
            type_shop = 'Yêu thích'
        elif 'YeGYFd LIaN-a' in str(product):
            type_shop = 'Mall'
        elif 'YeGYFd _0-VFOk' in str(product):
            type_shop = 'Yêu thích +'
        
        # Get sold_quatity
        try:
            sold_quatity = product.find('div', class_ = 'r6HknA uEPGHT').get_text().strip()
        except:
            sold_quatity = ''
        
        # Get location
        location = product.find('div', class_='zGGwiV').get_text().strip()
        
        # Get discount
        try:
            discount = product.find('div', class_='GOgNtl').get_text().strip()
        except:
            discount = ''
        
        # Get freeship 
        freeship = 'None'
        if '_8-xLHM' in str(product):
            freeship = 'Yes'
            
        # Get rating
        try:
            shop_rating = product.find_all('div', class_ = 'shopee-rating-stars__stars')
            rating_star = 0
            for shop in shop_rating:
                ratinglo = shop.find_all('div', class_ = 'shopee-rating-stars__lit')
        
                
                for xx in ratinglo:
                
                    percent = ''.join(re.findall(r'\d',xx['style']))
                
                    if(int(percent) == 100):
                        rating_star += 1
                    else:
                        rating_star += round(float(percent)/pow(10, len(percent)),1)
        except:
            rating_star = 'No reviews'
        
        # A product
        product_details = [products_name,type_shop,sold_quatity,location,discount,freeship,rating_star]
        
        # Start write to sheet
        write_to_sheet(product_details)
    
        
# Page by page   
def page_by_page(urlContinue):
    
        print(f'Start with {name}')
        print('...........................................')
        print('...........................................')
        
        for i in range(0,number_page_scan):
            index = i
            urlContinue = urlContinue + '&page=' + str(i)
            
            print(f'Start with {name} page {index + 1}')
            print('...........................................')
            
            
            driver.get(urlContinue)
            time.sleep(5)
            
            
            total_height = int(driver.execute_script("return document.body.scrollHeight"))
        
            for i in range(1, total_height, 5):
                driver.execute_script("window.scrollTo(0, {});".format(i))
                
            soup = BeautifulSoup(driver.page_source, 'html.parser')
    
            page_source = soup.select('.VTjd7p.whIxGK')
            
            
            scan_web(page_source) 

# Run each product
for name in listProduct:
    
    nameProduct = name.split(" ")
    urlProduct = "%20".join(nameProduct)
    urlContinue = url + urlProduct
    
    page_by_page(urlContinue)
    time.sleep(10)
      


# Complete
sheet.update('A1:G1',[title])
sheet.format('A1:G1', {'textFormat':{'bold':True}})

driver.quit()

    
        

