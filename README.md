# ItsyBitsyScript
## Introduction
This Python script automates the process of fetching stock data from TradingView and applying various filters to identify stocks meeting specific criteria. The script utilizes the Selenium web automation library to interact with the TradingView website, extracts relevant stock information, and updates an Excel spreadsheet with the results.

## Features
Fetching Stock Data: Retrieves stock names, market capitalization, trading volume, price-to-earnings ratio, and current price from TradingView.
Applying Filters: Implements filters based on market capitalization and simple moving averages to identify stocks meeting specified criteria.
Updating Excel Spreadsheet: Updates an Excel spreadsheet ('Stocks.xlsx') with the fetched data and filter results.
## Prerequisites
- Python 3.x

- Selenium library

- Openpyxl library

- Chrome WebDriver

## Installation
- Install Python

- Install required libraries:
pip install selenium openpyxl

- Download Chrome WebDriver: ChromeDriver Downloads

- Update the webdriverpath variable in the script with the path to your Chrome WebDriver executable.

## Usage
- Run the script:
python script_name.py

- The script will open the TradingView website, apply filters, fetch stock data, and update the 'Stocks.xlsx' Excel spreadsheet.
## Configuration
- Modify the Gmail credentials (gmail_user and gmail_password) in the sendmail function to set up email notifications.
- Adjust the file paths in the addDirect and addYesNo functions based on your project structure.
## Disclaimer
This script is for educational and informational purposes only. Use the data and results at your own risk. The developer is not responsible for any financial decisions made based on this script.
