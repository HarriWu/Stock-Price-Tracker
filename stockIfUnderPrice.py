#! /usr/bin/env python3
import random

import requests, bs4, smtplib, time, openpyxl


def compare_price(URL, set_price, user_agent, body):
    """ Compares set price for company stock to the price scraped off of Yahoo Finance
        using BeautifulSoup and returns an updated report.

    Args:
        URL: Yahoo Finance URL where price is scrapped.
        set_price: Desired price for that stock.
        user_agent: The User-Agent associated with computer.
        body: total report on which prices have dipped below desired prices.

    Returns:
        Updated total report on which prices have dipped below desired prices.
        body += '\nPrice right: ' + URL for success,
        body for otherwise.
    """
    headers = {
        "User-Agent": user_agent}

    res = requests.get(URL, headers=headers)

    if res:
        print('Successfully retrieved information')
        soup = bs4.BeautifulSoup(res.content, features="html.parser")
        data_extracted = soup.findAll('span', class_='Trsdu(0.3s) Fw(b) Fz(36px) Mb(-4px) D(ib)')
        price = data_extracted[0].getText()
        float_price = float(price.replace(',', ''))

        if set_price >= float_price:
            body += '\nPrice right: ' + URL
            return body
        else:
            return body

    else:
        print('Not Found')
        return body


def send_email(gmail_address, gmail_password,
               receiving_email, body):
    """ Sends email from a gmail account to a receiving email address.
    Args:
        gmail_address: sender gmail address.
        gmail_password: sender gmail password.
        receiving_email: receiving email address.
        body: the body of the email message being sent.
    """
    # Setting up sender email
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(gmail_address, gmail_password)

    subject = 'Price is under your set price'
    msg = f"Subject: {subject}\n\n{body}"

    # Sender email sends message to receiver email
    server.sendmail(
        gmail_address,
        receiving_email,
        msg
    )


def extracting_values(excel_file):
    """ Extracts all data from excel file and uses that data to compare
        prices to build a report that is sent by email.

    Args:
        excel_file: data.xlsx which contains all required data
                    regarding emails, User-Agent, URLs, and set prices.
    """
    wb = openpyxl.load_workbook(excel_file)
    sheet = wb['Sheet1']

    gmail_address = sheet.cell(row=1, column=2).value
    gmail_password = sheet.cell(row=2, column=2).value
    receiving_email = sheet.cell(row=3, column=2).value
    user_agent = sheet.cell(row=4, column=2).value

    index = 2
    body = ''

    # Excel file extraction so that URLs and set prices are able to be extracted
    while sheet.cell(row=index, column=3).value or sheet.cell(row=index, column=4).value is not None:
        URL = sheet.cell(row=index, column=3).value
        set_price = sheet.cell(row=index, column=4).value
        body = compare_price(URL, set_price, user_agent, body)
        print(body)
        # It checks everything at random so that you won't get blocked
        time.sleep(random.randint(10, 20))
        index += 1

    print('done')
    # Send email at the end of the day containing all the links to the stocks
    # that are under or equal to their desired price

    if body != '':
        send_email(gmail_address, gmail_password,
                   receiving_email, body)
    else:
        send_email(gmail_address, gmail_password,
                   receiving_email, 'No prices have lowered past your set prices')


# Daily checks on the prices
while (True):
    extracting_values('data.xlsx')
    time.sleep(86400)
