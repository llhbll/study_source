from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, StaleElementReferenceException
from openpyxl import Workbook, load_workbook
import time

wb  = Workbook()
sheet1 = wb.active
sheet1.title = '표준교육과정 과목중심' #시트명
sheet1.cell(row=1, column=2).value = "과목명"
sheet1.cell(row=1, column=3).value = '전공구분'
sheet1.cell(row=1, column=4).value = '강의시간'
sheet1.cell(row=1, column=5).value = '실습시간'

driver = webdriver.Chrome()

kconti_url = 'https://www.cb.or.kr/creditbank/stdPro/nStdPro1_1.do'
driver.get(kconti_url)
driver.maximize_window()

row = 2

def work(): # 전공하나하나 과목 얻어오기
    global row
    all_list = driver.find_element_by_css_selector('div.listDateWrap01')
    select_list = all_list.find_elements_by_css_selector('li')
    for item in select_list:
        jungong_flag_s = item.find_element_by_css_selector('em').text
        subject = item.find_element_by_css_selector('a').text
        lecture_time = item.find_elements_by_css_selector('span')[2].text
        practice_time = item.find_elements_by_css_selector('span')[3].text

        sheet1.cell(row=row, column=1).value = major_name
        sheet1.cell(row=row, column=2).value = subject
        sheet1.cell(row=row, column=3).value = jungong_flag_s
        sheet1.cell(row=row, column=4).value = lecture_time
        sheet1.cell(row=row, column=5).value = practice_time
        row = row + 1

    driver.back()
    time.sleep(2)

for seq in range(2,10, 2):
    form = "#contents > div.innerContView > div.stdProtResult > div > ul:nth-child({}) > li > a"
    url = form.format(str(seq))
    try: # 학사단위당 전공이 하나일 경우
        search_button = driver.find_element_by_css_selector(url)
        major_name = search_button.text
        search_button.click()
        work()
        continue
    except NoSuchElementException:  # 학사단위당 전공이 여러개일 경우는 예외상황 발생되므로 ...^^ 꼼수
        for i in 30: #학사단위당 전공이 기껏해봐야 30개 이내겠지?
            form = "#contents > div.innerContView > div.stdProtResult > div > ul:nth-child({}) > li:nth-child({}) > a"
            url = form.format(str(seq), str(i))
            try:
                search_button = driver.find_element_by_css_selector(url)
                major_name = search_button.text
                search_button.click()
                work()
                continue
            except NoSuchElementException:  # 학사단위당 전공 모두 작업끝내고 이후 시도할 경우
                break

driver.quit()

wb.save("./excel_folder/" + "모든전공.xlsx")
