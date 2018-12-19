import requests
from bs4 import BeautifulSoup
from datetime import date
from datetime import datetime, timedelta
import time
import io
import unicodedata
import pandas as pd
from itertools import chain

REPORT_FORMAT = "png"


class InvestmentReport:
    def __init__(self):
        pass

    def url_builder(self):
        """method prepares two url's of current date and last week's 
                nasdaqomxbaltic trading session prices. Url's will be used for scraping"""
        template_url1 = (
            "http://www.nasdaqbaltic.com/market/?pg=mainlist&date=yyyy.mm.dd&lang=en"
        )
        date_of_today = str(date.today()).replace(
            "-", "."
        )  # pulling current date in format of 'yyy.mm.dd'
        last_week_date = str(date.today() - timedelta(7)).replace(
            "-", "."
        )  # pulling last week's date in 'yyy.mm.dd'
        # two output URL's:
        self.url_prices_today = template_url1.replace("yyyy.mm.dd", date_of_today)
        self.url_prices_last_week = template_url1.replace("yyyy.mm.dd", last_week_date)


    def scrape_last_prices(self, url):
        """Method scrapes nasdaqomxbaltic.com prices list url with certain date and outputs csv file of listed companies last prices"""
        r = requests.get(url)
        soup = BeautifulSoup(r.text, "lxml")
        # finding all table rows:
        rows_containers = soup.find_all("tr")
        # extracting rows that actually carry information on listed companies trading data:
        companies_rows = []
        for i in rows_containers:
            if i.has_attr("id"):
                companies_rows.append(i)
        # preparing output list:
        self.stock_closing_prices = []

        for company in companies_rows):
            # temporary list for one company to be filled and appended to stock_closing_prices of lists.
            one_company_info = ["", "", ""]
            td_tag_within_company_row = company.findAll("td")
            # company name for i item
            company_name = td_tag_within_company_row[0].text

            # company ticker for i item
            ticker = td_tag_within_company_row[1].text
            company_ticker = ticker.replace("\xa0", "")

            # company last price for i item
            company_last_price = td_tag_within_company_row[4].text

            # filling temporary list
            one_company_info[0] = company_name
            one_company_info[1] = company_ticker
            one_company_info[2] = company_last_price

            # making list of lists:
            self.stock_closing_prices.append(one_company_info)

        # using pandas to export list of lists to a csv file:
        self.my_df = pd.DataFrame(self.stock_closing_prices)
        # export_filenaname = 'nasdaq'+str(url)
        self.my_df.to_csv(
            "nasdaq_prices"+url[-18:]+".csv", index=False, header=False
        )


    def performace_evaluation(self):
        '''method takes two csv files, and returns one excel file with two sheets of top and worst 5 (10 in total)
        performers. Method also returns two lists for potential further scraping of related companies announcements'''
        # reading csv data:
        last_week_df = pd.read_csv('nasdaq_prices2018.12.11&lang=en.csv', names=['Company', 'Ticker', 'Last week price'])
        today_df = pd.read_csv('nasdaq_prices2018.12.18&lang=en.csv', names=['Company', 'Ticker', 'Last price'])
        data_from_last_week_df = last_week_df[['Ticker', 'Last week price']]
        data_from_today_df = today_df['Last price']
       
        # joining dataframes and making a relative price change calulation:
        joint_df = pd.concat([data_from_last_week_df, data_from_today_df], axis=1, join='inner')
        joint_df['Change, %'] = (joint_df['Last price'] - joint_df['Last week price'])/joint_df['Last week price']
        
        # top and worst performers from sorted joint dataframe:
        top_performers = joint_df.sort_values(by='Change, %', ascending=False).drop(columns = ['Last week price', 'Last price']).head(5)
        worst_performers = joint_df.sort_values(by='Change, %').drop(columns = ['Last week price', 'Last price']).head(5)
        
        # writing to a common excel file, different sheets:
        writer = pd.ExcelWriter('Performance.xlsx')
        top_performers.to_excel(writer, sheet_name='Top performers', index=False)
        worst_performers.to_excel(writer, sheet_name='Worst Performers', index=False)
        writer.save()

        #creating two lists for potential future scraping:
        top_performers_to_be_list = top_performers.drop(columns = ['Change, %'])
        worst_performers_to_be_list = worst_performers.drop(columns = ['Change, %'])
        self.top_performers_list = list(chain.from_iterable(top_performers_to_be_list.values.tolist()))
        self.worst_performers_list = list(chain.from_iterable(worst_performers_to_be_list.values.tolist()))


    def run(self):
        self.url_builder()
        self.scrape_last_prices(self.url_prices_last_week)
        self.scrape_last_prices(self.url_prices_today)
        self.performace_evaluation()
        #self.generate_report()
        # self.get_twitter()
        # url.builder()
        # scrape_last_prices( url1 ...)
        # generate_report
        # other methods required to generate a report


if __name__ == "__main__":
    gimme_report = InvestmentReport()
    gimme_report.run()


