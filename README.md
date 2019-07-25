# Weekly Investment Report

## Description
Program generates automated weekly content beneficial to investors of Baltic Region (Nasdaq OMX Baltic stock market).

Program combines data from two sources:
- [Nasdaqomxbaltic](http://www.nasdaqomxbaltic.com) website;
- personal excel "database"-screener named "OMXB analitika.xlsm"

scrapes new prices and edits a copy of provided screener.
Resulting values are pushed to prepared excel template (data/Template.xlsx).

Final output is a PNG image ready to be shared in social networks. See examples below.

## Output Examples

### Example #1

![Report 2019.07.25](data/Report%202019.07.25.png)

### Example #2

To be added another day

## Requirements

- Python 3.7.3 +
- Modules (and versions) listed in requirements.txt
- Child `/data` folder containing:
    - Template.xlsx
    - vbaProject.bin

### Workaround Screener issue

User will not be able to replicate the same output, because `screener_handler.py` module refers to `OMXB analitika.xlsm` and manipulates it's contents, which is held in author's computer and contains semi-sensitive information. To by-pass this, in `report_generator.py` `run` method disable (delete/comment out) this line:

- `self.get_data_from_screener(self.today_csv_filename)` - which queries screener file and returns a list of outputs.

Additionally user should disable (delete/comment out) `report_generator.py` `load_data_to_template_excel` method lines:
- `ws['A21'] = self.random_ratio`
- `ws['B22'] = self.random_bool`
- `self.df_from_screener.to_excel(writer, index=False, header = False, startrow = 23, startcol = 20,  sheet_name = 'Main')
`

This way Program only updates the upper section of output image (scrapes prices, calculates price change, pushes to Template.xlsx)

## Installation

Simply execute report_generator.py according to requirements listed above.

### Code Structure

Program employs the benefits of Object Oriented Programming.
Two modules `report_generator.py` (main file) and `screener_handler.py` each have two classes `InvestmentReport` and `Screener` accordingly.

Each module/class has a `run` method, which is a combination of class methods when run as standalone. Methods' docstrings are self explanatory in most cases

## Potential Future Development:

- Migrating excel screener to database solution for full-blown web solution/app to access a ton of information;
- Twitter bot
- Setting up program on a server and send email/SMS whenever new report has been generated (each Saturday morning i.e.)

### Acknowledgements

- [Aidis Stukas](https://github.com/aidiss) - my first Python bootcamp instructor. 
- Robertas Skalskas for guidelines on joining two modules