import datetime
import logging
import smtplib
import sys
import time
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from logging.handlers import RotatingFileHandler

import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By

# 로그 설정
logger = logging.getLogger('[NHIS]annMonitoring')
logger.setLevel(logging.DEBUG)

# 로그 포맷 설정
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')

# 로그 파일 핸들러 설정
file_handler = RotatingFileHandler('monitoring.log', maxBytes=1024 * 1024, backupCount=5, encoding='utf-8')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(formatter)

# 콘솔 핸들러 설정
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.DEBUG)
console_handler.setFormatter(formatter)

# 로거에 핸들러 추가
logger.addHandler(file_handler)
logger.addHandler(console_handler)

# 키워드
keyWord = '중단'

# 메일 정보
sender = 'shfin098@gmail.com'
password = 'iqbivkukjxwzhlwu'
receiver_email = ['shp5314@naver.com', 'destiny0909@naver.com']
alarm_email = ['shp5314@naver.com', 'destiny0909@naver.com', 'jshan@fintech1.co.kr', 'tmkim@fintech1.co.kr']
test_email = ['shp5314@naver.com']

# 현재 시간
now = datetime.datetime.now()
minute = now.minute
nowFm = now.strftime('%Y-%m-%d %H:%M:%S')

# 발견 메일
def send_mail(sender, password, alarm_email, content):
    # SMTP 이메일
    message = MIMEMultipart()
    message['From'] = '건강보험공단'
    message['to'] = ", ".join(alarm_email)
    message['subject'] = '!!!!!!![NHIS] 점검 공지사항 모니터링   ' + now.strftime('%Y%m%d') + '!!!!!!!'
    message.attach(MIMEText(content, 'plain'))

    # SMTP 서버 연결
    smtp_server = smtplib.SMTP("smtp.gmail.com", 587)
    smtp_server.starttls()
    smtp_server.login(sender, password)

    # 이메일 전송
    smtp_server.sendmail(sender, alarm_email, message.as_string())

    # SMTP 서버 연결 종료
    smtp_server.quit()

# 확인 메일
def do_mail(sender, password, test_eamil, content):
    # SMTP 이메일
    message = MIMEMultipart()
    message['From'] = '건강보험공단'
    message['to'] = ", ".join(test_eamil)
    message['subject'] = '[NHIS] 점검 공지사항 모니터링   do_' + now.strftime('%Y%m%d')
    message.attach(MIMEText(content, 'plain'))

    # SMTP 서버 연결
    smtp_server = smtplib.SMTP("smtp.gmail.com", 587)
    smtp_server.starttls()
    smtp_server.login(sender, password)

    # 이메일 전송
    smtp_server.sendmail(sender, alarm_email, message.as_string())

    # SMTP 서버 연결 종료
    smtp_server.quit()

# 게시글 필터링
def ann_alarm():
    try:
        url = 'https://www.nhis.or.kr/nhis/index.do'
        options = webdriver.ChromeOptions()
        options.add_argument('headless')
        options.add_argument('--window-size=1024,768')
        options.add_argument('--disable-gpu')
        driver = webdriver.Chrome('/Users/csw/AppData/Local/Programs/Python/Python310/chromedriver.exe', options=options)
        driver.get(url)
        time.sleep(3)

        driver.find_element(By.XPATH, '//*[@id="main-wrap"]/div[4]/div[1]/div[1]/ul/li[1]/a').click()
        time.sleep(2)

        ann_table = driver.find_elements(By.TAG_NAME, 'table')

        elements = pd.read_html(ann_table[0].get_attribute('outerHTML'))

        dataFrame = elements[0]
        dataFrame.to_excel('nhis_announcement.xlsx')

        df = pd.read_excel('nhis_announcement.xlsx')

        now = datetime.datetime.now()
        now_dt = now.strftime('%Y.%m.%d')

        title_df = pd.read_csv('title_list.txt')
        list = title_df['0'].tolist()

        titles = []
        for i in range(len(df)):
            title = df.iloc[i]['제목']
            reg_dt = df.iloc[i]['등록일']
            view_cnt = df.iloc[i]['조회수']

            if reg_dt == now_dt and view_cnt < 3 and title not in list :
                content = nowFm + '_new_ann_title : ' + title
                if keyWord in title:
                    content =  nowFm +'_inspection_ann_title : ' + title
                    logger.info('키워드 알림 (title : ' + title + ')')
                    send_mail(sender, password, alarm_email, content)
                else:
                    send_mail(sender, password, alarm_email, content)
                    logger.info('새 공지사항 알림 (title : ' + title + ')')
            titles.append(title)

        # 제목을 CSV 파일로 저장
        df_titles = pd.DataFrame(titles)
        df_titles.to_csv('title_list.txt', index=False)


    except Exception as e:
        logger.error(str(e))
        logger.error(traceback.format_exc())


ann_alarm()

try:
    if minute == 0:
        log_msg = f'[{nowFm}] 프로그램이 실행 중 입니다.'
        content = log_msg
        logger.info(log_msg)
        do_mail(sender, password, receiver_email, content)
except Exception as e:
    logger.error(str(e))
    logger.error(traceback.format_exc())
