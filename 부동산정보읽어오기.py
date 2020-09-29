import re
import numpy as np
import pandas as pd
import requests
from datetime import datetime
from bs4 import BeautifulSoup


def get_naver_realasset(area_code, page=1):
    url = 'http://land.naver.com/article/articleList.nhn?' \
          + 'rletTypeCd=A01&tradeTypeCd=A1&hscpTypeCd=A01%3AA03%3AA04' \
          + '&cortarNo=' + area_code \
          + '&page=' + str(page)
0
    r = requests.get(url)
    soup = BeautifulSoup(r.text, 'lxml')

    table = soup.find('table')
    trs = table.find_all('tr')
    if '등록된 매물이 없습니다' in trs[0].text:
        return pd.DataFrame()

    value_list = []

    # 거래, 종류, 확인일자, 매물명, 면적(㎡), 층, 매물가(만원), 연락처
    for tr in trs[::2]:
        tds = tr.find_all('td')
        cols = [' '.join(td.text.strip().split()) for td in tds]

        if '_thumb_image' not in tds[3]['class']:  # 현장확인 날짜와 이미지가 없는 행
            cols.insert(3, '')

        # print(cols)
        거래 = cols[0]
        종류 = cols[1]
        확인일자 = datetime.strptime(cols[2], '%y.%m.%d.')
        현장확인 = cols[3]
        매물명 = cols[4]
        면적 = cols[5]
        공급면적 = re.findall('공급면적(.*?)㎡', 면적)[0].replace(',', '')
        전용면적 = re.findall('전용면적(.*?)㎡', 면적)[0].replace(',', '')
        공급면적 = float(공급면적)
        전용면적 = float(전용면적)
        층 = cols[6]
        if cols[7].find('호가일뿐 실거래가로확인된 금액이 아닙니다') >= 0:
            pass  # 단순호가 별도 처리하고자 하면 내용 추가
        매물가 = int(cols[7].split(' ')[0].replace(',', ''))
        연락처 = cols[8]

        value_list.append([거래, 종류, 확인일자, 현장확인, 매물명, 공급면적, 전용면적, 층, 매물가, 연락처])

    cols = ['거래', '종류', '확인일자', '현장확인', '매물명', '공급면적', '전용면적', '층', '매물가', '연락처']
    df = pd.DataFrame(value_list, columns=cols)
    return df

df = get_naver_realasset('1168010600')

print(df)