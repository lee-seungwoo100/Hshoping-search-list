
import time
import re
from datetime import datetime
from selenium import webdriver
import requests , openpyxl
from bs4 import BeautifulSoup
from selenium.webdriver.common.keys import Keys


excel_file = openpyxl.Workbook()

search_product ='고구마'

options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument("disable-gpu")
options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/99.0.4844.84 Safari/537.36')

browser = webdriver.Chrome(options=options) #'./chromdriver.exe'


# 홈앤쇼핑 특정키워드로 셀레니움 정보창 정보 가져오기

browser.get('https://www.hnsmall.com/search/search.do?query_top=%EA%B3%A0%EA%B5%AC%EB%A7%88')
browser.implicitly_wait(10)
browser.find_element_by_name('query_top').send_keys(search_product)
browser.find_element_by_id('tcampaign').click()
time.sleep(3)
H_source = browser.page_source

# 홈앤쇼핑 정보창 토대로 순위내용 가져오기 및 엑셀정리 /인기상품순 정렬
excel_sheet2 = excel_file.active
excel_sheet2.append(['순위','상품명', '가격','최종혜택가', '리뷰수','판매수','링크'])
excel_sheet2.title = '홈앤쇼핑'
excel_sheet2.column_dimensions['A'].width = 5
excel_sheet2.column_dimensions['B'].width = 80
excel_sheet2.column_dimensions['C'].width = 15
excel_sheet2.column_dimensions['D'].width = 15
excel_sheet2.column_dimensions['E'].width = 10
excel_sheet2.column_dimensions['F'].width = 15

soup = BeautifulSoup(H_source,'html.parser')
data = soup.select('div.viewTypeImg.v2')
for item in data:
    item_all = item.select('ul > li ')
    for index, item_1 in enumerate(item_all,start=1):
        item_link = item_1.select_one('a')['href']
        item_name = item_1.select('div.textZone')
        item_price = item_1.select('div.commentZone')

        for name in item_name:
            name_1 = name.select_one('p.goodsName > a > span').get_text().strip()
            try:
                price_1 = name.select_one('p.sale_price > em').get_text().strip()   
            except:
                price_1 = name.select_one('p.price > em').get_text().strip()
            try:
                price_2 = name.select_one('p.benefit_price > em').get_text().strip()
            except:
                price_2 = "혜택가 없음"
        for price in item_price:
            rate = price.select_one('span.cmt_count')
            if rate:
                rate = price.select_one('span.cmt_count').get_text().strip()[1:-2]
            elif rate ==None:
                rate = '리뷰없음'
            review = price.select_one('span.cmt_sales').get_text().strip()
            if re.search('판매량',review):
                review = price.select_one('span.cmt_sales').get_text().strip()
            else:
                review = ''
            print(index,name_1,price_1,price_2,rate,review,item_link)
            excel_sheet2.append([index,name_1,price_1,price_2,rate,review,item_link])
            excel_sheet2.cell(row=index, column=7).hyperlink = item_link

cell_A1 = excel_sheet2['A1'] # 셀 선택하기
cell_A1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_A1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_B1 = excel_sheet2['B1'] # 셀 선택하기
cell_B1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_B1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_C1 = excel_sheet2['C1'] # 셀 선택하기
cell_C1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_C1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_D1 = excel_sheet2['D1'] # 셀 선택하기
cell_D1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_D1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_E1 = excel_sheet2['E1'] # 셀 선택하기
cell_E1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_E1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors

cell_F1 = excel_sheet2['F1'] # 셀 선택하기
cell_F1.alignment = openpyxl.styles.Alignment(horizontal='center') # 중앙정렬하기
cell_F1.font = openpyxl.styles.Font(color="01579B") # 폰트 색깔 바꾸기
# 색상값 찾기: https://material.io/design/color/#tools-for-picking-colors            

browser.quit()

excel_file.save(search_product+'베스트'+'.xlsx')
excel_file.close()
