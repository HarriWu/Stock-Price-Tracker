# Stock Price Tracker

App checks stock price on Yahoo Finance and if that price falls below or equals your desired price for that stock an email will be sent as a reminder. The app uses an excel file to hold all of the stock URLs and your set price for them. Each stock will be checked and a daily report will be sent via email of the stock prices that have met the desired price that you set.

## How It Works

Uses openpyxl to extract all information regarding sender email, receiver email, user-agent, stock URLs and their desired prices from excel file. Uses BeautifulSoup to scrape Yahoo Finance for the stock price. Compiles report based on comparing stock price with desired price and sends out report daily to the receiver email address.

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes. See deployment for notes on how to deploy the project on a live system. 

### Installation

Install all the third-party modules required in the directory where data.xlsx and stockIfUnderPrice.py are located.


Install requests
```
$ sudo pip3 install requests
```
Install BeautifulSoup
```
$ sudo pip3 install beautifulsoup4
```
Install openpyxl
```
$ sudo pip3 install openpyxl
```

## Setup

In data.xlsx for sender gmail address and password enter in a gmail account that will send the email to your receiver email address.
To fill out the User-Agent part -> Search google.
```
my user agent 
```
Put the Yahoo Finance stock urls under the URLs column and put your desired price beside in the Set Price column.

## Deployment

Navigate to the directory where data.xlsx and stockIfUnderPrice.py are located in Terminal and change the .py file’s permissions to make it executable

```
$ chmod +x stockIfUnderPrice.py
```

To run

```
$ ./stockIfUnderPrice.py
```

## Dependencies

* BeautifulSoup - Scrapping library
* Requests - Python HTTP client interface library
* openpyxl - Python library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files






