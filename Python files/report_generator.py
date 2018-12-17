import requests
from bs4 import BeautifulSoup
from datetime import date
import time
import io
import unicodedata
import pandas as pd

REPORT_FORMAT = 'png'

class InvestmentReport:

        def __init__(self):
                pass

        def url_builder(self):
                '''method inserts current date into nasdaq's url, which will be ready to be used for scraping'''
                template_url1 = 'http://www.nasdaqbaltic.com/market/?pg=mainlist&date=dd.mm.yyyy&lang=en'
                date_of_today = str(date.today()).replace('-','.') #pulling current date
                year = date_of_today[0:4]  #extracting year
                date_for_scrape_1 = date_of_today[5:] + '.' + year   #building suitable string
                url_today_prices = template_url1.replace('dd.mm.yyyy', date_for_scrape_1)

        def scrape_last_prices(self, url)
                '''scrape_last_prices method outputs csv file of Nasdaqomxbaltic last prices'''
                r = requests.get(InvestmentReport.url_builder())
                soup = BeautifulSoup(r.text, "lxml")
                #finding all table rows:
                rows_containers = soup.find_all("tr")
                #extracting rows that actually carry information on listed companies trading data:
                companies_rows=[]
                for i in rows_containers:
                if i.has_attr('id'):
                        companies_rows.append(i)
                #preparing output list:
                stock_closing_prices=[]

                for i in range(len(companies_rows)):
                #temporary list for one company to be filled and appended to stock_closing_prices of lists.
                one_company_info=['','','']
                td_tag_within_company_row = companies_rows[i].findAll('td')
                #company name for i item
                company_name = td_tag_within_company_row[0].text    
                        
                #company ticker for i item
                ticker = td_tag_within_company_row[1].text
                company_ticker = ticker.replace('\xa0', '')
                        
                #company last price for i item
                company_last_price = td_tag_within_company_row[4].text

                # filling temporary list
                one_company_info[0]=company_name
                one_company_info[1]=company_ticker
                one_company_info[2]=company_last_price

                #making list of lists:    
                stock_closing_prices.append(one_company_info)

                #using pandas to export list of lists to a csv file:
                my_df = pd.DataFrame(stock_closing_prices)

                my_df.to_csv('nasdaq_last_prices_export_list.csv', index=False, header=False)




