import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import logging
import csv
import os

logging.basicConfig(level=logging.DEBUG)

REPORT_FORMAT = "png"

def extract_company_name(td_tag_within_company_row):
    company_name = td_tag_within_company_row[0].text
    return company_name

def extract_company_ticker(td_tag_within_company_row):
    ticker = td_tag_within_company_row[1].text
    ticker = ticker.replace("\xa0", "")
    return ticker

def extract_last_price(td_tag_within_company_row):
    last_price = td_tag_within_company_row[4].text
    return last_price

class InvestmentReport:
    def __init__(self):
        pass

    def server_response_checker(self):
        '''method returns server response code of nasdaqomxbaltic.com website.'''
        nasdaqomxbaltic_home_url = 'http://www.nasdaqbaltic.com/market/?lang=en'
        r_check = requests.get(nasdaqomxbaltic_home_url)
        self.server_response = r_check.status_code
    
    def url_builder(self):
        """method prepares two url's of current date and last week's 
        nasdaqomxbaltic trading session prices. Url's will be used for scraping"""
        template_url1 = ("http://www.nasdaqbaltic.com/market/?pg=mainlist&date=yyyy.mm.dd&lang=en")
        # pulling dates for template url of today and last week:
        self.date_of_today_string = datetime.now().strftime("%Y.%m.%d")
        self.last_week_date_string = (datetime.now() - timedelta(7)).strftime("%Y.%m.%d")
        # two output URL's:
        self.url_prices_today = template_url1.replace("yyyy.mm.dd", self.date_of_today_string)
        self.url_prices_last_week = template_url1.replace("yyyy.mm.dd", self.last_week_date_string)

    def get_prices_soup(self, url):
        '''Creates a soup from url'''
        r = requests.get(url)
        self.soup = BeautifulSoup(r.text, features='lxml')

    def get_rows_containing_data(self):
        '''Extracts rows containing actual companies data'''
        rows_containers = self.soup.find_all("tr")
        self.companies_rows = []
        for i in rows_containers:
            if i.has_attr("id"):
                self.companies_rows.append(i)
        
    def get_scrape_results(self):
        '''Iterates through row elements and outputs a list of dictionaries for each company.'''
        #General list of dictionaries for all companies:
        self.scrape_output = []
        #iterating through companies to collect data:
        for company in self.companies_rows:
            temp_d = {}
            td_tag_within_company_row = company.findAll("td")
            temp_d['Name'] = extract_company_name(td_tag_within_company_row)
            temp_d['Ticker'] = extract_company_ticker(td_tag_within_company_row)
            temp_d['Last Price'] = extract_last_price(td_tag_within_company_row)
            # output list of dictionaries:
            self.scrape_output.append(temp_d)

    def csv_filename(self, url):
        '''Prepares appropriate csv file names, depending on what url was passed'''
        self.today_csv_filename = 'data/Prices_' + self.date_of_today_string + '.csv'
        self.last_week_csv_filename = 'data/Prices_' + self.last_week_date_string + '.csv'
        #csv filename depends on URL being used for scraping:
        if self.date_of_today_string in url:
            self.temp_csv_filename = self.today_csv_filename
        elif self.last_week_date_string in url:
            self.temp_csv_filename = self.last_week_csv_filename
        else:
            logging.debug("Faulty URL passed to function")

    def export_scrape_results_csv(self):
        keys = self.scrape_output[0].keys()
        with open (self.temp_csv_filename, 'w', newline='') as outputfile:
            dict_writer = csv.DictWriter(outputfile, keys)
            dict_writer.writeheader()
            dict_writer.writerows(self.scrape_output)

    def scrape_to_csv(self, url):
        '''List of methods to be performed to scrape and output desired data to csv'''
        logging.debug('Executing scrape to csv(url)')
        self.get_prices_soup(url)
        self.get_rows_containing_data()
        self.get_scrape_results()
        self.csv_filename(url)
        self.export_scrape_results_csv()
        logging.debug('Scrape and output of method scrape_to_csv(url) successfully executed')

    def form_joint_dataframe(self):
        '''Takes two csv files, cleans data and merges to joint dataframe. Calculates additional column ['Change, %']'''
        last_week_df = pd.read_csv(self.last_week_csv_filename, header = 0, encoding = "ISO-8859-1", names=['Name', 'Ticker', 'Last Week Price'])
        today_df = pd.read_csv(self.today_csv_filename, header = 0, encoding = "ISO-8859-1", names=['Name', 'Ticker', 'Last Price'])
        #trimming dataframes:
        data_from_last_week_df = last_week_df[['Ticker', 'Last Week Price']]
        data_from_today_df = today_df[['Ticker', 'Last Price']]
        # joining dataframes and making a relative price change calulation:
        self.joint_df = pd.merge(data_from_last_week_df, data_from_today_df, on='Ticker')
        self.joint_df['Change, %'] = (self.joint_df['Last Price'] - self.joint_df['Last Week Price'])/self.joint_df['Last Week Price']
        
    def df_to_excel(self):
        '''Adjusts joint dataframe from form_join_dataframe(); Forms two dataframes and outputs them to Peformance.xlsx
        that cointains two sheets: Top Performers & Worst Performers. Additionally - deletes unneccessary csv files'''
        # top and worst performers from sorted merged dataframe:
        top_performers = self.joint_df.sort_values(by='Change, %', ascending=False).drop(columns = ['Last Week Price', 'Last Price']).head(5)
        worst_performers = self.joint_df.sort_values(by='Change, %').drop(columns = ['Last Week Price', 'Last Price']).head(5)
        # writing results a common excel file, different sheets:
        writer = pd.ExcelWriter('data/Performance.xlsx')
        top_performers.to_excel(writer, sheet_name='Top Performers', index=False)
        worst_performers.to_excel(writer, sheet_name='Worst Performers', index=False)
        writer.save()
        #delete csv files:
        os.remove(self.today_csv_filename)
        os.remove(self.last_week_csv_filename)

    def run(self):
        self.server_response_checker()
        if self.server_response == 200:
            self.url_builder()
            logging.debug("URL's prepared")
            logging.debug("About to scrape this URL: " + self.url_prices_today)
            self.scrape_to_csv(self.url_prices_today)
            logging.debug("About to scrape another URL: " + self.url_prices_last_week)
            self.scrape_to_csv(self.url_prices_last_week)
            logging.debug("2 csv files were created in /data folder")
            self.form_joint_dataframe()
            logging.debug("Joint dataframe was created")
            self.df_to_excel()
            logging.debug("data/Performance.xlsx has been created; temporary csv files have been deleted")
        else:
            logging.warning('\nWebsite is currently unreachable;\nProgram has terminated.')


if __name__ == "__main__":
    gimme_report = InvestmentReport()
    gimme_report.run()