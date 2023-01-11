# version 1.1.0

import multiprocessing
from datetime import timezone, timedelta, date, time, datetime
import requests
from bs4 import BeautifulSoup
import pandas as pd

url = 'https://search.naver.com/search.naver'


def clpr_crawling(it):
    it = it.split('\n')
    if len(it) >= 3:
        it = ' '.join(it[0:-2])
    else:
        it = ' '.join(it)

    params = {
        'query': f'{it}'
    }

    response = requests.get(url, params=params)

    if response.status_code != 200:
        print(response.status_code)
        return ['-', '-', '-', '-']

    html = response.text
    soup = BeautifulSoup(html, 'html.parser')

    try:
        href = soup.select_one('#_cs_root > div.ar_spot > div > h3 > a').attrs['href']
    except:
        print(f'href error with {it}')
        return ['-', '-', '-', '-']

    response = requests.get(href)

    if response.status_code != 200:
        print(f'{response.status_code} error with {it}')
        return ['-', '-', '-', '-']

    html = response.text
    soup = BeautifulSoup(html, 'html.parser')

    try:
        price = soup.select_one(
            '#content > div.section.invest_trend > div.sub_section.right > table > tbody > tr:nth-child(2) > td:nth-child(2) > em').text.strip()
        if price == '':
            price = 0
    except:
        try:
            price = soup.select_one(
                '#chart_area > div.rate_info > div > p.no_today > em > span:nth-child(1)').text.strip()
        except:
            price = '-'

    try:
        foreigner = soup.select_one(
            '#content > div.section.invest_trend > div.sub_section.right > table > tbody > tr:nth-child(2) > td:nth-child(4) > em').text.strip().strip(
            '+')
        if foreigner == '':
            foreigner = 0
    except:
        foreigner = '-'

    try:
        institution = soup.select_one(
            '#content > div.section.invest_trend > div.sub_section.right > table > tbody > tr:nth-child(2) > td:nth-child(5) > em').text.strip().strip(
            '+')
        if institution == '':
            institution = 0
    except:
        institution = '-'

    try:
        BA = soup.select_one('#tab_con1 > div:nth-child(3) > table > tr.strong > td > em').text.strip().strip('%')
        if BA == '':
            BA = 0
    except:
        print('B/A error with ' + it)
        BA = '-'

    return [price, foreigner, institution, BA]


if __name__ == '__main__':
    today = date.today()

    multiprocessing.freeze_support()

    df = pd.read_excel('./종목 리스트.xlsx', sheet_name='Sheet1', engine='openpyxl')
    df = df.dropna(axis=0)
    print(df)

    ri = pd.RangeIndex(start=1, stop=len(df.index) + 1, step=1)
    df.index = ri

    print(ri)

    interests = df.iloc[:, 0]

    interests = interests.tolist()

    pool = multiprocessing.Pool(processes=8)
    data_list = pool.map(clpr_crawling, interests)

    columns = [f'{today.month}/{today.day} 종가', '외국인', '기관', '외국인소진율']

    print('Collected Price Info')

    df = pd.concat([df.iloc[:, 0], pd.DataFrame(data_list, columns=columns, index=ri)],
                   axis=1)
    df.to_excel(f'./종가{today.month}-{today.day}.xlsx', sheet_name=f'{today.month}-{today.day}')

    print('Excel Generated Successfully!')
