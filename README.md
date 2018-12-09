# Weekly Investment Report

## Trumpas aprašymas:

Weekly automatinis report script-generatorius su paruoštu stiliaus template (galbūt koks nors API, galbūt PY biblioteka) PDF formatu bei PNG formatais (content fb page/ twitter paskyroms)


## Duomenų šaltiniai:

- [Nasdaqomxbaltic](http://www.nasdaqomxbaltic.com)
- excel failas-duombazė


## Report turinys:

- Top 5 best/worst performers of the week. Pvz:

![pavyzdys](https://github.com/yomajo/myproject/blob/master/Images/pvz%20performers.JPG?raw=true "Pavyzdys")

- Scrape related news from these performers.
- savaitės uždarymo kainas supushinti į excel duombazę, ištraukti išrūšiuotas top 10 pagal pigumą (tam tikras pagal kainas apskaičiuotas rodiklis excel duombazėje)
- pateikti selected santykinius rodiklius su naujausiomis penktadienio sesijos uždarymo kainomis.
- įkalt disclaimer apačioj (static text)

## Potencialūs tobulinimai:
- twitter boto kūrimas
- ir šio reporto tweetinimas penktadienio vakarą


## Kriterijai projektui:

- naudingumas
- originalumas
- testavimas
- stilius
- dokumentacija
- versijų kontrolė
- 'import this'

Now get to `coding`.

___

## Program structure (*unfinished*)

### stock_info_scraper.py Method "stock_closing_prices"

Class employing BeautifulSoup library. Scrapes 

Inputs:
- initiation time
- date for trading session we want the data from

Output: 
- three lists inside the list:

`[ [company1, company2, ...], [company_ticker1, company_ticker2, ...], [closing_price1, closing_price2] ]`

- handle list and export as csv/excel

Source url:  http://www.nasdaqbaltic.com/market/?pg=mainlist&date=07.12.2018&lang=en

### performance_evaluation.py

Handles output from stock_closing_prices, compares to according list of previous period (week) and outputs two arrays of companies of best and worst performance over the last week.

Input:
- output from stock_closing_prices over two dates passed by datehander.py

Output:
- csv/excel two arrays of best and worst performing stocks over the last week.


### datehander.py

datehander.py script tracks current time, puts it into perspective of current date. If current time is Friday 18:00, it calls for action, outputs two dates.

Input:
- current time

Output:
- handles initiation and excecusion of main project script that creates automated report.
- current friday date
- last friday date


### stock_info_scraper.py Method "related_annoncements"
Another Scraper Class method should have 10 companies as inputs from performance_evaluation.py module
and scrape related news from these companies over the period (starting and ending date range) again defined
by datehhandler.py 

Inputs:
- 10 companies from performance_evaluation.py
- date range from datehandler.py

Output:
- Hyperlink with Title as text and url to actual announcement on nasdaqomxbaltic.com

Source url: http://www.nasdaqbaltic.com/market/?page=1&issuer=&market=&legal%5B0%5D=main&legal%5B1%5D=firstnorth&start=2018-11-01&end=2018-12-07&keyword=&pg=news&lang=en&currency=EUR&downloadcsv=0 

### Screener (excel file) screener_handler.py (*unfinished*)
Script takes closing prices of all companies, inputs into certain sheet, cells new prices, sorts table by certain criteria and outputs 5 companies based on it.

### report_generator.py (*unfinished*)
Takes all the required inputs, template excel file, fills the fields, and saves a new document in multiple formats.