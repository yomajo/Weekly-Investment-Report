import requests
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import pandas as pd
import logging
import csv
import os
import openpyxl
import excel2img

#LOGGING SETTINGS:
logging.basicConfig(level=logging.DEBUG)

#GLOBAL VARIABLES
TEMPLATE_PATH = 'data/Template.xlsx'
REPORT_FORMAT = '.png'

#PUBLIC FUNCTIONS USED IN SCRAPPER
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
        '''Extracts rows containing actual companies data (containing attribute selector "id" in HTML)'''
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
        
    def get_best_worst_performers_df(self):
        '''Forms two dataframes of best and worst companies by price performance, sorted accordingly'''
        self.top_performers = self.joint_df.sort_values(by='Change, %', ascending=False).drop(columns = ['Last Week Price', 'Last Price']).head(5)
        self.worst_performers = self.joint_df.sort_values(by='Change, %').drop(columns = ['Last Week Price', 'Last Price']).head(5)

    def Load_data_to_template_excel(self):
        '''Loads dataframes, variables to pre-made Template.xlsx, without modifying the rest
        of the document. Saves changes'''
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        ws = wb['Main']
        ws['B18'] = self.date_of_today_string
        ws['B19'] = self.last_week_date_string
        with pd.ExcelWriter(TEMPLATE_PATH, engine='openpyxl') as writer:
            writer.book = wb
            writer.sheets = dict((ws.title, ws) for ws in wb.worksheets)
            self.top_performers.to_excel(writer, index=False, header = False, startrow = 1, sheet_name = 'Main')
            self.worst_performers.to_excel(writer, index=False, header = False, startrow = 7, sheet_name = 'Main')
            writer.save()
    
    def clean_temp_files(self):
        '''Cleans unneccessary files after program has finished'''
        os.remove(self.today_csv_filename)
        os.remove(self.last_week_csv_filename)

    def generate_output(self):
        '''Generates desired image file from Template\'s named range'''
        self.report_file_name = 'Report '+self.date_of_today_string + REPORT_FORMAT
        excel2img.export_img(TEMPLATE_PATH, 'data/' + self.report_file_name, None, 'Output_Area')


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
            self.get_best_worst_performers_df()
            logging.debug("Two dataframes of best and worst performing stocks formed")
            self.Load_data_to_template_excel()
            logging.debug("All desired data loaded to Template.xlsx")
            self.clean_temp_files()
            logging.debug("Temporary csv files have been deleted")
            self.generate_output()
            logging.debug('Report named: "'+ self.report_file_name + '" is created in data/ folder')
        else:
            logging.warning('\nWebsite is currently unreachable;\nProgram has terminated.')


if __name__ == "__main__":
    gimme_report = InvestmentReport()
    gimme_report.run()