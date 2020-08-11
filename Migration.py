# -*- coding: utf-8 -*-
"""
Created on Sun Jun 21 23:39:39 2020

@author: Pandey
"""

from selenium import webdriver
import pandas as pd
import time
from selenium.webdriver.common.keys import Keys
from tqdm import tqdm, trange
import numpy as np
import datetime as dt

from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from bs4 import BeautifulSoup
from selenium.webdriver.support.select import Select
import sys


path="C:\\Users\\OYO\\Desktop\\migrated files\\migrated_list.xlsx"
df=pd.read_excel(path)
list_of_acco_ids = list(df['Acco-Ids'].unique())
#last_ten_acco_ids = list_of_acco_ids[-11:]


driver = webdriver.Chrome(executable_path=r"C:\Users\OYO\Desktop\chromedriver.exe")
driver.maximize_window()

driver.get("https://www.oyo-vacation-homes.net/cgi/lars/algm/login.htm")
driver.find_element_by_xpath('//*[@id="loginid"]').send_keys('')
driver.find_element_by_xpath('//*[@id="passwd"]').send_keys('')
driver.find_element_by_xpath('//*[@id="signin"]').click()

driver.find_element_by_xpath('//*[@id="quicksearchmenu"]').send_keys(Keys.CONTROL, 'a')

#list_of_acco_ids = ['AT-6561-27','IT-60010-081']
#tkey=list_of_acco_ids[1]
FinalData={}
idschecked=[]
p2021= '//*[contains(concat( " ", @class, " " ), concat( " ", "showPlanNextYear", " " ))]'

def callfunc(p):
    i=p
    print(i)
    def fetchdata(idstobechecked,i):
        for tkey in idstobechecked:
            idschecked.append(tkey)

            driver.find_element_by_xpath('//*[@id="quicksearchmenu"]').send_keys(Keys.CONTROL, 'a')
            driver.find_element_by_xpath('//*[@id="quicksearchmenu"]').send_keys(tkey)
            driver.find_element_by_xpath('//*[@id="quicksearchmenu"]').send_keys(Keys.RETURN)
            time.sleep(10)
            try:
                
                if(i==1):
                    driver.find_element_by_xpath(p2021).click() #to see 2021
                    time.sleep(5)
            except:
                print("No plandays")

            days = driver.find_elements_by_class_name('planday')

            try:

                starts = pd.DataFrame([x for x in days if 'href="javascript:;"' in x.get_attribute('innerHTML') and 'font color="FFFFFF"' in x.get_attribute('innerHTML') and 'class="calender-a-tag plnItem_HPBZ"' in x.get_attribute('innerHTML')] ,columns = ['Day Elements'])
                FinalData[tkey]=[tkey,"NA",0,0]

                starts['HTML'] = starts.apply(lambda x: x['Day Elements'].get_attribute('outerHTML'), axis = 1)



                datadate=[]

                for i in starts['HTML']:
                    #print(i)
                    indexdd = i.find('data-date')
                    dd=i[indexdd+11:indexdd+19]
                    datadate.append(dd)
                starts['data-date']=datadate


                pricelist=[]
                for d in starts['Day Elements']:

                    d.click()
                    time.sleep(3)

                    driver.find_element_by_xpath('//*[@id="plnItem_body"]/button').click() #to see type
                    time.sleep(5)

                    select = Select(driver.find_element_by_xpath('//select[@id="bookingOrClosed"]'))
                    selected_option = select.first_selected_option
                    #print (selected_option.text)

                    if('migrated' in (selected_option.text).lower()):#=='Homeowner booking - Migrated (Bookings for other guests and personal stay above limit)'):
                        #print("True")
                        price_element=driver.find_element_by_xpath('//input[@id="hoBookingAmount"]')
                        price=price_element.get_attribute("value")
                        pricelist.append(float(price.replace(",",".")))
                        #print(price)

                    elif('migrated' not in (selected_option.text).lower()):#!='Homeowner booking - Migrated (Bookings for other guests and personal stay above limit)'):
                        driver.find_element_by_xpath('//*[@id="show_plnItem_modal"]/div/div/div[3]/button[4]').click() #to close
                        pricelist.append(0)
                        time.sleep(5)
                        continue



                    try:
                        driver.find_element_by_xpath('//*[@id="show_plnItem_modal"]/div/div/div[3]/button[4]').click() #to close
                        time.sleep(5)
                    except:
                        pass



                starts['Price_List']=pricelist #THIS IS NOT WORKING IF EVEN 1 HOME IS PERSONAL
                sumofpricelist=sum(pricelist)



                FinalData[tkey]=[tkey,datadate,pricelist,sumofpricelist]
            except:
                print('no days')

            return idschecked


    allkeys = list_of_acco_ids #l
    len_ids = len(list_of_acco_ids)

    while(len(idschecked) != len_ids):
        idstobechecked = list(set(allkeys) - set(idschecked))
        fetchdata(idstobechecked,i)
        print("ids remaining: ",idstobechecked)

    #new addition
    df=pd.DataFrame.from_dict(FinalData)

    if len(FinalData) is not 0:

        Finaldf=pd.DataFrame()
        acco_ids=df.loc[0]
        amount=df.loc[3]
        datadates = df.loc[1]
        acco_ids.reset_index().drop('index',axis=1)
        amount.reset_index().drop('index',axis=1)
        Finaldf['acco_ids']=acco_ids
        Finaldf['Amount']=amount
        Finaldf['DataDates']=datadates
        Finaldf.reset_index(drop=True,inplace=True)

    print("THE END")
    sys.stdout.write('\a')
    sys.stdout.flush()
    idschecked.clear()
    return Finaldf

df2020 = callfunc(0)
df2021 = callfunc(1)

df2020.rename(columns={'Amount':'2020GBV','DataDates':'DD2020'},inplace=True)
df2021.rename(columns={'Amount':'2021GBV','DataDates':'DD2021'},inplace=True)
fdf = df2020.merge(df2021, left_on='acco_ids', right_on='acco_ids')

fdf['GBV Sum'] = fdf['2020GBV']+fdf['2021GBV']

path = 'C:/Users/OYO/Desktop/migrated files/migrated_gbv - 28July_original.xlsx'
fdf.to_excel(path)
