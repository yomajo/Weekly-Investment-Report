# Weekly Investment Report

Scope of program exceeds the term of the Python Programming Course, therefore this desription is devided into "Current Program", "Future Development" and "Potential Program Extensions" sections.

In short: Program should generate automated weekly content beneficial to investors of Baltic Region (Nasdaq OMX Baltic stock market). Report should be an excel template saved each week as a PNG or PDF file.

## Current and future data sources:

- [Nasdaqomxbaltic](http://www.nasdaqomxbaltic.com)
- excel database 'screener.xlsx' 

## Current Program:

### Installation
Simply execute report_generator.py from any IDE. Be sure you have installed modules listed in requirements.txt

### Code Structure

Program employs the benefits of Object Oriented Programming.
Currently one class InvestmentReport is sufficient to cover the needs of the program.

Class methods:

- url_builder: takes template url ("http://www.nasdaqbaltic.com/market/?pg=mainlist&date=yyyy.mm.dd&lang=en") and creates two url's based on date. One Url is for prices of today's trading session (last prices), while the other, needed for comparisson url is from last week session.
- server_response_checker: simply checks if nasdaqomxbaltic website is responding with 200 code. If not - program terminates.
- scrape_last_prices: scrapes listed companies prices and outputs a csv file for the url passed in.
- performance_evaluation: takes two csv files from different time-based url scrapes and outputs excel file with two sheets, that contains top and worst performers. Method also outputs the list of these companies tickers for future code implementations.
- run: method simply combines program excecusion sequence. Method also has condition, and if server is not responding, further methods are not executed and logging.warning is retured to console.

## Future Development

- Add excel screener handler. Enable rewrite of last prices, extraction of selected sorted metrics from one sheet.
- create report_template.xlsx
- insert top/worst performers, selected data from screener, other important metrics to report, plot graphs, output a PNG/PDF.

## Potential Program Extensions:
- Twitter bot

## Criteria:

- usability
- originality
- testing
- style
- documentation
- version control
- 'import this'

Now get to `coding`.