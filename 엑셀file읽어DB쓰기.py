from selenium import webdriver
import requests, re
from bs4 import BeautifulSoup
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
import time
from openpyxl import Workbook, load_workbook
import pyperclip
from selenium.webdriver.common.keys import Keys
import sys

import sqlite3

con = sqlite3.connect('./hakjum.db')
cur = con.cursor()

#처음만 해줌
#cur.execute('CREATE TABLE t_major(Haksa_n, Major_n, Subject_n, jungong_f, lecture_tm, pratice_tm);')
#cur.execute('INSERT INTO PhoneBook VALUES("Lim ChanHyuk", "010-8443-8473");')
#cur.execute('SELECT * FROM PhoneBook;')


load_wb  = load_workbook("./excel_folder/모든전공.xlsx", data_only=True) #data_only 수식이 아닌 값으로 읽음
load_ws = load_wb['표준교육과정 과목중심']
for row in load_ws.rows:
    row_list = []
    for cell in row:
        row_list.append(cell.value)
    cur.execute('INSERT INTO t_major VALUES(row_list[0], row_list[1], row_list[2], row_list[3], row_list[4],'
                'row_list[5]);')

con.commit()
con.close()
        #item.find_element_by_css_selector('a').click()  #가격을 구하기 위해서






