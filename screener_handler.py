import logging
import os
from shutil import copy
import openpyxl
import csv
import random 

#LOGGING SETTINGS:
logging.basicConfig(level=logging.DEBUG)

#GLOBAL VARIABLES
#Note, that due to sensitive information inside screener it is not publicly available
SCREENER_ABS_PATH = 'C:/Svarbu/OMXB analitika.xlsm'
SCREENER_FILENAME = SCREENER_ABS_PATH.split('/')[-1]
SRC_DIR = '/'.join(SCREENER_ABS_PATH.split('/')[:-1])

class Screener:
    def __init__(self):
        pass
    
    def screener_exists(self, abs_path):
        '''Checks if screener file exists and returns True if it does. Takes one arg - absolute path to file of interest'''
        if os.path.exists(abs_path):
            return True

    def make_temp_screener_copy(self, src_path):
        '''Creates a temporary copy of file to project /data folder. Takes an input of source path as string'''
        copy(src_path, os.getcwd() + '/data/')

    def last_price_csv_to_dict(self, csvfilename):
        '''Takes argument of csv filename as a string. Method creates dictionary of company ticker as key and last price as value'''
        with open('data/' + csvfilename) as csv_file:
            csv_reader = csv.reader(csv_file)
            #skipping header row
            next (csv_reader)
            self.ticker_price_dict={}
            for csv_line in csv_reader:
                try:
                    self.ticker_price_dict[csv_line[1]] = float(csv_line[2])
                except:
                    self.ticker_price_dict[csv_line[1]] = csv_line[2]
    
    def open_screener(self):
        '''Trying to manipulate fresh copy of xlsm file'''
        logging.debug('Opening worksheet {}'.format('data/' + SCREENER_FILENAME))
        # Limitation of openpyxl about loosing formulas or being unable to read cell contents where formulas are present
        # Workaround: open two sessions in different modes; one for reading, another for writing
        self.wb = openpyxl.load_workbook('data/' + SCREENER_FILENAME)
        self.wb_read = openpyxl.load_workbook('data/' + SCREENER_FILENAME, data_only=True)

    def load_new_prices(self):
        '''Editing inside copy of excel screener'''
        ws = self.wb['Prices']
        ws_read = self.wb_read['Prices']
        #Determining at which row should price uploading begin:
        for cell in ws['E']:
            if cell.value == 'Last':
                self.input_range_row = cell.row
                logging.debug(f'Price uploading should start at row number: {self.input_range_row}')
                break

        self.count_uploaded = 0
        #iterating excel ticker range, if a match in csv formed dict is found - value is updated
        for cell in range(self.input_range_row, ws.max_row):
            if ws_read[f'W{cell}'].value in self.ticker_price_dict:
                ws[f'E{cell}'].value = self.ticker_price_dict[ws_read[f'W{cell}'].value]
            self.count_uploaded += 1
        logging.debug(f'A total number of {self.count_uploaded} companies last prices were uploaded to temp. excel file')

    def use_latest_values(self):
        '''Sets additional cell values, so most up to date data is used in screener file'''
        ws = self.wb['Summary']
        ws['C1'] = self.wb['Universals']['E18'].value
        ws['C2'] = ws['X1'].value
        logging.debug('Additinal cell values have been set in Summary sheet C1:C2 with values {} and {}'.format(self.wb['Universals']['E18'].value, ws['X1'].value))

    def get_available_ratios(self):
        '''Get a collection of ratios available in the screener'''
        ws = self.wb_read['Summary']
        self.ratios_dict = {}
        # Starting location of ratios and their categories in Summary sheet defined by start_row_col tuple
        start_row_col = (4, 29)
        #Iterating through ratio categories (row=4), and ratios (columns 29-31) to collect a dict of lists for each category
        logging.debug('Looping through [Summary] sheet and collecting available ratios')        
        for col in range(start_row_col[1], ws.max_column):
                if ws.cell(row=start_row_col[0], column=col).value != None:
                        temp_key = ws.cell(row=start_row_col[0], column=col).value
                        ratios_within_category = []
                        # If outter column (ratio category) cell value is not emply - start another cycle of collecting ratios iterating through rows
                        for row in range(start_row_col[0] + 1, ws.max_row):
                                if ws.cell(row=row, column=col).value != None:
                                        ratios_within_category.append(ws.cell(row=row, column=col).value)
                                else:
                                        self.ratios_dict[temp_key] = ratios_within_category
                                        break             
                else:
                        break

    def pick_random_values_for_sorting(self):
        '''outputs three randomly picked values: 1. ratio category 2. ratio 3. True/False boolean for accending/descending sorting'''
        self.random_ratio_category = random.choice(list(self.ratios_dict))
        self.random_ratio = random.choice(self.ratios_dict[self.random_ratio_category])
        self.random_boolean = random.choice([True, False])

    def close_screener(self):
        '''saves a separate copy of screener and closes workbook'''
        self.wb.save('data/' + 'OMXB analitika.xlsx')
        self.wb.close()
        self.wb_read.close()
        logging.debug('All of the above operations performed are visible in newly created: {}'.format('OMXB analitika.xlsx'))
        logging.debug('Finished')


    def run(self):
        if self.screener_exists(SCREENER_ABS_PATH) == True:
            logging.debug('Screener {} found'.format(SCREENER_FILENAME))
            self.make_temp_screener_copy(SCREENER_ABS_PATH)
            logging.debug('Screener {} copy created in {} folder'.format(SCREENER_FILENAME, '/data'))
            # 
            self.open_screener()
            self.last_price_csv_to_dict('Prices.csv')
            self.load_new_prices()
            self.use_latest_values()
            self.get_available_ratios()
            self.pick_random_values_for_sorting()

            self.close_screener()
            # 
        else:
            logging.warning(' PROGRAM TERMINATED. \n Please check directory: {} \n File named {} was not found there.'.format(SRC_DIR, SCREENER_FILENAME))
        
if __name__ == '__main__':
    Screener().run()