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


def return_soup(url):
    """ returns HTML data in python readable way """
    macro_text = requests.get(f'{url}').text
    soup = BeautifulSoup(macro_text, 'lxml')
    return soup


def format_list(start, u_list, integer=3):
    """creates list with sub-lists of each year's ratios"""
    if start == integer:
        year_ratios = [u_list[x:x+4] for x in range(0, len(u_list), 4)]
    else:
        year_ratios = [u_list[x:x+4] for x in range(start + 1, len(u_list), 4)]
        year_ratios.insert(0, u_list[0:start + 1])
    return year_ratios


def range_dict(dates, u_list):
    """ creates dictionary with year as key and min to max range as value """
    range_d = {}

    for index, ratio_year in enumerate(u_list):
        if len(ratio_year) <= 1:
            range_d[int(dates[0]) - index] = ratio_year
        else:
            range_d[int(dates[0]) - index] = [min(ratio_year), max(ratio_year)]

    return range_d


def year_start(dates):
    """ Returns starting point for the ratio list """
    date_str = ' '.join(dates)
    return date_str.rindex(dates[0]) // 5


def ratio_range_s(url, flag=False):
    """ Creates ratio range dict with single table layout """
    soup = return_soup(url)

    data_table = soup.find('table', class_='table')
    data = data_table.find_all('td')

    ratios = []
    dates = []
    ttm_fcf = []
    # splits the date from yyyy-mm-dd format into just the year and
    # only appends every 4th item in list (which is date)
    # appends every 3rd item (ratio) to list ratios
    for index, item in enumerate(data):
        dates.append(str(item.text.split('-')[0])) if index % 4 == 0 else None
        ratios.append(float(item.text)) if index % 4 == 3 else None
        if flag:
            if index == 2 and item.text != '' or index == 6:
                ttm_fcf.append(float(item.text.strip('$')))

    start = year_start(dates)
    year_ratios = format_list(start, ratios)
    range_d = range_dict(dates, year_ratios)

    if flag:
        return range_d, ttm_fcf[0]

    return range_d


def ratio_range_d(url):
    """ Creates ratio dict with dual table layout """
    soup = return_soup(url)

    table = soup.find_all('div', class_='col-xs-6')
    quarter_table = table[1]
    data = quarter_table.find_all('td')

    eps = []
    dates = []

    for index, item in enumerate(data):
        dates.append(str(item.text.split('-')[0])) if index % 2 == 0 else None
        eps.append(float(item.text.strip('$') if item.text.find(',') == -1 else item.text.replace(',', '').strip('$'))) \
            if index % 2 == 1 else None

    start = year_start(dates)
    year_eps = format_list(start, eps)
    range_d = range_dict(dates, year_eps)

    return range_d


if __name__ == '__main__':
    u = input('Enter the url: ')
    s = ratio_range_s(u)
    print(s)
