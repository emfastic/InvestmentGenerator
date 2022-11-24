""" Used to verify float for numeric attributes and to initialize dict as a default value"""
from dataclasses import dataclass, field
import numpy as np
from scraping_macrotrends import ratio_range_s, ratio_range_d
from scraping_yf import current_stock_data


@dataclass
class Stock:
    """ Creates a stock object with the following attributes comp name, ticker,
        current share price, current forward PE, beta, current PEG, current enterprise ratio,
         eps range that MacroTrends has, current FCF ratio, FCF range that MacroTrends has,
         current PB ratio, and PB ratio range that MacroTrends has """
    company_name: str = 'unknown'
    ticker: str = 'unknown'
    current_price: float = 0.0
    forward_pe: float = 0.0
    beta: float = 0.0
    PEG: float = 0.0
    enterprise_ratio: float = 0.0
    eps_range: dict = field(default_factory=dict)
    current_fcf_ratio: float = 0.0
    fcf_range: dict = field(default_factory=dict)
    current_pb: float = 0.0
    pb_range: dict = field(default_factory=dict)


def create_stock(ticker, company_name):
    """ Uses Yahoo Finance and MacroTrends scrapers to create a stock with all attributes"""
    stock = Stock()
    ticker = ticker.upper()
    url_company_name = company_name.lower().replace(' ', '-')
    yf_url = f'https://finance.yahoo.com/quote/{ticker}/key-statistics?p={ticker}'
    mt_eps_url = f'https://www.macrotrends.net/stocks/charts/{ticker}/{url_company_name}/eps-earnings-per-share-diluted'
    mt_fcf_url = f'https://www.macrotrends.net/stocks/charts/{ticker}/{url_company_name}/price-fcf'
    mt_pb_url = f'https://www.macrotrends.net/stocks/charts/{ticker}/{url_company_name}/price-book'

    yf_dict = current_stock_data(yf_url)
    eps_dict = ratio_range_d(mt_eps_url)
    fcf_dict, ttm_fcf = ratio_range_s(mt_fcf_url, flag=True)
    pb_dict = ratio_range_s(mt_pb_url)

    stock.company_name = company_name
    stock.ticker = ticker
    stock.current_price = yf_dict['current price']
    stock.forward_pe = yf_dict['forward P/E']
    stock.beta = yf_dict['beta']
    stock.PEG = yf_dict['PEG']
    stock.enterprise_ratio = yf_dict['enterprise ratio']
    stock.eps_range = eps_dict

    if ttm_fcf not in ('N/A', 0):
        stock.current_fcf_ratio = float(np.round(stock.current_price / ttm_fcf, 2))
    else:
        stock.current_fcf_ratio = 'N/A'

    stock.fcf_range = fcf_dict

    if yf_dict['book val per share mrq'] != 'N/A' and yf_dict['book val per share mrq'] != 0:
        stock.current_pb = float(np.round(stock.current_price / yf_dict['book val per share mrq'], 2))
    else:
        stock.current_pb = 'N/A'

    stock.pb_range = pb_dict

    return stock


if __name__ == '__main__':
    tick = input('Enter the ticker symbol: ')
    co = input('Enter the company name: ')
    s = create_stock(tick, co)
