""" bs4 to process the HTML data that requests receives from website """
from bs4 import BeautifulSoup
import os
import sys


def override_where():
    """ overrides certifi.core.where to return actual location of cacert.pem"""
    # change this to match the location of cacert.pem
    print(os.path.abspath("cacert.pem"))
    return os.path.abspath("cacert.pem")


# is the program compiled?
if hasattr(sys, "frozen"):
    import certifi.core

    os.environ["REQUESTS_CA_BUNDLE"] = override_where()
    certifi.core.where = override_where

    # delay importing until after where() has been replaced
    import requests
    import requests.utils

    # replace these variables in case these modules were
    # imported before we replaced certifi.core.where
    requests.utils.DEFAULT_CA_BUNDLE_PATH = override_where()


def current_stock_data(url):
    """ retrieves and returns dict of following
        stock data from YahooFinance statistics page
        forward P/E, PEG, enterprise ratio, current price, beta,
        and book val from mrq """
    macro_text = requests.get(f'{url}',
                              headers={'User-Agent': 'Custom'}).text
    soup = BeautifulSoup(macro_text, 'lxml')

    # retrieves the current data from the Yahoo Finance table on statistics page
    full_stats = soup.find_all('td', {'class': 'Ta(c) Pstart(10px) Miw(60px) Miw(80px)--pnclg '
                                               'Bgc($lv1BgColor) fi-row:h_Bgc($hoverBgColor)'})

    relevant_stats = {'forward P/E': full_stats[3].text, 'PEG': full_stats[4].text,
                      'enterprise ratio': full_stats[8].text}

    current_price = soup.find('fin-streamer', {'class': 'Fw(b) Fz(36px) Mb(-4px) D(ib)'}).text
    relevant_stats['current price'] = current_price

    beta = soup.find('td', {'class': 'Fw(500) Ta(end) Pstart(10px) Miw(60px)'}).text
    relevant_stats['beta'] = beta
    # iterate through each table (all have same class) return the 8th in index
    # which is balance sheet
    balance_sheet = soup.find_all('div', {'class': 'Pos(r) Mt(10px)'})[8]
    # iterate through each numeric value in the balance sheet and
    # return the book val
    book_val_mrq = balance_sheet.find_all('td', {'class': 'Fw(500) '
                                                          'Ta(end) Pstart(10px) Miw(60px)'})[5].text
    relevant_stats['book val per share mrq'] = book_val_mrq

    for key, value in relevant_stats.items():
        if value != 'N/A':
            relevant_stats[key] = float(value)
        else:
            continue

    return relevant_stats


if __name__ == '__main__':
    u = input('Enter the url: ')
    print(current_stock_data(u))
