import requests
from bs4 import BeautifulSoup
from datetime import date
from datetime import datetime, timedelta
import time
import io
import unicodedata
import pandas as pd

REPORT_FORMAT = 'png'

class InvestmentReport:

        def __init__(self):
                pass

        def url_builder(self):
                '''method prepares two url's of current date and last week's 
                nasdaqomxbaltic trading session prices. Url's will be used for scraping'''
                template_url1 = 'http://www.nasdaqbaltic.com/market/?pg=mainlist&date=yyyy.mm.dd&lang=en'
                date_of_today = str(date.today()).replace('-','.') #pulling current date in format of 'yyy.mm.dd'
                last_week_date = str(date.today()- timedelta(7)).replace('-','.') #pulling last week's date in 'yyy.mm.dd'
                #two output URL's:
                self.url_prices_today = template_url1.replace('yyyy.mm.dd', date_of_today)
                self.url_prices_last_week = template_url1.replace('yyyy.mm.dd', last_week_date)

                print(self.url_prices_today)
                print(self.url_prices_last_week)


        def scrape_last_prices(self, url):
                '''scrape_last_prices method outputs csv file of Nasdaqomxbaltic last prices'''
                r = requests.get(self.url_prices_today)
                soup = BeautifulSoup(r.text, "lxml")
                #finding all table rows:
                rows_containers = soup.find_all("tr")
                #extracting rows that actually carry information on listed companies trading data:
                companies_rows = []
                for i in self.rows_containers:
                        if i.has_attr('id'):
                                companies_rows.append(i)
                #preparing output list:
                stock_closing_prices=[]

                for i in range(len(companies_rows)):
                #temporary list for one company to be filled and appended to stock_closing_prices of lists.
                        one_company_info = ['','','']
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




        def run(self):
                self.url_builder()
                
                # self.get_twitter()
                # url.builder()
                # scrape_last_prices( url1 ...)
                # generate_report
                # other methods required to generate a report


if __name__ == '__main__':
        gimme_report = InvestmentReport()
        gimme_report.run()


# my_dates = InvestmentReport()
# my__test_output = my_dates.url_builder()


   




