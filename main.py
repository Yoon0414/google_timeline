from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import shutil
import pyautogui
import re
import time
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta

ID = input('구글 ID를 입력하시오')
PW = input('구글 PW를 입력하시오')

while True:
    start = input('탐색 시작 연도, 월, 일을 8자리로 입력하시오')
    end = input('탐색 종료 연도, 월, 일을 8자리로 입력하시오')
    if datetime.strptime(start, '%Y%m%d') <= datetime.strptime(end, '%Y%m%d'):
        break
    else:
        print("날짜를 다시 입력해주세요")
        # 시작 날짜 > 종료 날짜일 때

year = datetime.strptime(start, '%Y%m%d').year
month = datetime.strptime(start, '%Y%m%d').month
day = datetime.strptime(start, '%Y%m%d').day
# YMD 정보로 치환

day_len = (datetime.strptime(end, '%Y%m%d') - datetime.strptime(start, '%Y%m%d')).days + 1
# 날짜 차이 계산

google_dict = {'day': [], 'walk': [], 'car': [], 'subway': [], 'bus': []}


walk_xpath = '//*[@data-activity="2"]'  # 도보 xpath
car_xpath = '//*[@data-activity="29"]'  # 자가용 xpath
subway_xpath = '//*[@data-activity="9"]'  # 지하철 xpath
bus_xpath = '//*[@data-activity="7"]'  # 버스 xpath

try:
    shutil.rmtree(r"C:\chrometemp")  # 브라우저에 저장된 정보 지우기
except FileNotFoundError:
    pass

subprocess.Popen(r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222 '
                 r'--user-data-dir="C:\chrometemp"')
option = Options()
option.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
# 디버거 크롬 열기

driver = webdriver.Chrome(r'C:\Users\이윤병\Downloads\chromedriver\chromedriver.exe', options=option)
# 크롬 웹드라이버 주소 기입

driver.implicitly_wait(10)

driver.get('https://www.google.com/maps/timeline?ved=0ENaiAggAKAA&pb&gid=103209328611938709185&pli=1&rapt=AEjHL'
           '4Mkm8TI334zyzc44XzBsd8YIsWIet9Unoc-tQxJ17Z8PS_0FA3QSjqCVOvcv_UREt6krbJeR8fkRxtebhiqIXrySiW-NA')
# 구글 타임라인 주소로 이동

pyautogui.write(ID)  # 아이디
pyautogui.press('tab', presses=3)
pyautogui.press('enter')
time.sleep(3)
pyautogui.write(PW)  # 비밀번호
pyautogui.press('enter')
driver.implicitly_wait(10)
# --------------------------------------------------------------------------------------------------구글 타임라인 접속까지

driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/div[2]/div[1]/div[1]/div[1]').click()
pyautogui.press('down', presses=(datetime.today().year + 1 - year))
pyautogui.press('enter')  # 원하는 연도로 이동

time.sleep(1)

driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/div[2]/div[1]/div[2]/div[1]').click()
pyautogui.press('down', presses=month)
pyautogui.press('enter')  # 원하는 달로 이동

time.sleep(1)

aria_lebel = '//*[@aria-label=""]'
str_day = str(day) + ' ' + str(month) + '월'
aria_lebel = aria_lebel[:17] + str_day + aria_lebel[17:]
driver.find_element_by_xpath('/html/body/div[1]/div[6]/div[1]/div[2]/div[1]/div[3]/div[1]').click()
driver.find_element_by_xpath(aria_lebel).click()  # 원하는 날짜로 이동


def active(xpath):
    try:
        way = driver.find_element_by_xpath(xpath).text
        way_num = float(re.findall('\\d+.\\d+', way)[0])  # 정수.정수 꼴 찾기
        if 'km' not in way:
            way_num = way_num / 1000  # km 단위로 환산
    except:
        way_num = 0

    return way_num


for i in range(day_len):
    google_dict['walk'].append(active(walk_xpath))
    google_dict['car'].append(active(car_xpath))
    google_dict['subway'].append(active(subway_xpath))
    google_dict['bus'].append(active(bus_xpath))

    add_day = str(datetime.strptime(start, '%Y%m%d')+relativedelta(days=i))
    google_dict['day'].append(add_day[:10])
    try:
        driver.find_element_by_xpath('//*[@id="map-page"]/div[2]/div/div/div/div[1]/i[2]').click()  # 다음 날짜로 넘어가기
    except:
        break


def average_and_append(act, term):
    act_sum = round(sum(google_dict[act])/term, 2)
    google_dict[act].append(act_sum)
# 평균 계산 & 딕셔너리에 추가


google_dict['day'].append('평균')
average_and_append('walk', day_len)
average_and_append('car', day_len)
average_and_append('subway', day_len)
average_and_append('bus', day_len)

result = pd.DataFrame.from_dict(google_dict)
result.to_excel('google.xlsx')
# 엑셀로 저장
