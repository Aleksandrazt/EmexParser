import requests
import json
import pandas as pd
#  import openpyxl нужен будет для правильной работы pandas
from bs4 import BeautifulSoup


def check_access(url: str):
    """
    Check if this page exists
    :param url: url of site with needed article
    :return: if page exist
    """
    if requests.get(url):
        return True
    else:
        return False


def save_data(info: dict):
    """
    Save information to exel in data_res.xlsx
    :param info: dict with keys Артикль, Кросс-номер and Бренд and values as list of articles cross-nums and brands
    for Excel
    """
    data = pd.DataFrame.from_dict(info)
    writer = pd.ExcelWriter('data_res.xlsx')
    data.to_excel(writer, index=False)
    writer.save()


def get_goods_list(url: str):
    """
    Get information about goods with needed article from site
    :param url: url with article
    :return:
    """
    if check_access(url):
        r = requests.get(url, params={'compression': 'gzip-js', 'ip': '1', '_': '1668210589993'})
        soup = BeautifulSoup(r.content, 'html.parser')
        goods = soup.find('script', id="__NEXT_DATA__")
        goods = json.loads(goods.text)['props']['initialState']['details']['makes']['list']
        return goods
    else:
        print("Response Failed")


def get_brand_and_number(goods: list):
    brand_and_numbers = {}
    for product in goods:
        brand_and_numbers[product['make']] = product['num']
    return brand_and_numbers


def get_page(article: str):
    """
    Create url with article abd return lists of brands and numbers
    :param article: needed article
    :return: brand_and_numbers.keys(), brand_and_numbers.values() that is lists with needed info
    """
    url = f'https://emex.ru/products/{article}'
    goods = get_goods_list(url)
    brand_and_numbers = get_brand_and_number(goods)
    return brand_and_numbers.keys(), brand_and_numbers.values()


def main():
    """
    Read articles from Excel data.xlsx and find info for each of them
    """
    articles = pd.read_excel('data.xlsx')['Артикул']
    info = {'Артикул': [], 'Кросс-номер': [], 'Бренд': []}
    for article in articles.values:
        brands, nums = get_page(article)
        info['Артикул'].extend([article for _ in range(len(nums))])
        info['Кросс-номер'].extend(nums)
        info['Бренд'].extend(brands)
    save_data(info)


if __name__ == "__main__":
    main()
