import json
import pprint

import openpyxl
import requests
import time
import requests
import random
import datetime
import openpyxl
import pandas as pd
from pyautogui import size
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import subprocess
import shutil
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities
from bs4 import BeautifulSoup
import time
import datetime
import pyautogui
import pyperclip
import csv
import sys
import os
import math
import requests
import re
import random
import chromedriver_autoinstaller
from PyQt5.QtWidgets import QWidget, QApplication, QTreeView, QFileSystemModel, QVBoxLayout, QPushButton, QInputDialog, \
    QLineEdit, QMainWindow, QMessageBox, QFileDialog
from PyQt5.QtCore import QCoreApplication
from selenium.webdriver import ActionChains
from datetime import datetime, date, timedelta
import numpy
import datetime
from window import Ui_MainWindow
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *


def getUrl(keyword):
    dataList = []
    count = 1
    endFlag=False
    dataPrev={}
    sameCount=0
    while True:
        # if count>=2:
        #     break
        cookies = {
            'NNB': 'DSFIMBNTPBPGI',
            'nx_ssl': '2',
            'ASID': 'd3d7bf49000001881cf2740500000046',
            'nid_inf': '-1284035119',
            # 'NID_AUT': 'WuLSEC5nK+ib0TSryTGY6GlKj08nYB+0Nt1L4EutHlgWTzp+zIjQvuC7gyE6hol0',
            # 'NID_JKL': '9AmEcwfMQE0cYzXMkcf5fLLiUd+yxWR3svcyEyt8moY=',
            # 'CBI_SNS': 'naver|mSZUwX9yJa5EY6hd',
            # 'CBI_SES': 'X5tpj5h5j/50OmSRq3rC303XG/YlPaiAq9Dh0iB+8aEVtNwNfB2fEwbMp36LJCIsAjFiXqBR+ZcR6Fm0MqPInAuRN9j/k38Q+jP1MpGMz1JnZe8dbEGVSRvpBT6vXbJZe40tdCyd8tAWFz/shvS8xRV5p3kTWrGlpfaEmNnZEFfBbQq8D3V6wfpR1SChyerRt98IGH8qslI7tegndfK10f/+WgMiFqZxIsWj0XgGVyxho0tM8Q9vcAvvLgfx+dyudvnco1QDxE+ixJWJmsG1bP9g2XsQyEY+WxJoqBspFByXhqmU5PxEPTjP7CEyBCNkOSQK/AbbBzZ5XZJZWlDTcCaeP/ykmnKWk/kvDaZTmAH3adwO0Vb7cjyD6obOB2URByMgAF6tZ2Suep3J3jbxycknJk/zotAbjQ7bsKwDwXzuH7SzWE8yE/uBXiSDjuB2sRxFuEUjruj3jVNGlW79GA==',
            # 'CBI_CHK': '"r5V0mf9uRUZHZ/vmLGy3ez7f4/k4aqWXL5o03eN68foZy3oPS8xU0v9iEv36wMxiHNPFJKIqnS8jaj+67l0NWg4BY2+gOeAANxCkpUdZggtLPOganbD/HumMEE/+hyh5kYpO8HQWLy3pTXrffuPRqqh9YGJAbleDrO2flGDOKc0="',
            # 'NID_SES': 'AAABoBssrRTBsd6G0Ycqw/VBy0VWP1Yzfezk1ZK1rQk/7tt2R1YXYdgUwAt8guRUK+J/zToY6oJXZHLhBgStyB5K/K7uMULNCmbq50a1k7XuF1wtQNR3BGMcsGmMQlldFi0ORGahWPt1YJALKgtsqD9E0I6l7xU7P+BY5qToQEyl1YN0HZvxT3N0GL2pQV2mvrwiTetzeu32LEhSoUcqt1lrV16ZLN3lK8A4cileOfwfy9bdheFGSU8dstOA/7ndI0MPheCSe5MzTn1zR0/O0ApCJtqdETlylwJq68MvmghVfI/sL+vjQkFJScPKaxhCtnTLi8vFnbpixQg8hzISLHuUOlkfXwFeiNuPA6aftIzPKW3NDSWZhu8cPVHv2K8mBod9wjEB0XDeVgoSolCUJSGL60McAhhHzK983LxP4VqUhLJoyBep1CAHFXLK3WhjhKdvlATL72w88baVPSkMbCWUeNrt5wvPTtNkcoYGFpjEVAgAliJtzyjwG5MK5mSfvbYuu61KjZOHIOHlNeHrL2pUEBqwLFFERXZ4/KFK9eskkjgX',
            # 'page_uid': 'i4UdIwprvTossB36spsssssstdV-461708',
            # 'JSESSIONID': 'E7D53BE263AEF89A6E3A0FA122AA0C6E.jvm1',
        }

        headers = {
            'authority': 'section.blog.naver.com',
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
            # 'cookie': 'NNB=DSFIMBNTPBPGI; nx_ssl=2; ASID=d3d7bf49000001881cf2740500000046; nid_inf=-1284035119; NID_AUT=WuLSEC5nK+ib0TSryTGY6GlKj08nYB+0Nt1L4EutHlgWTzp+zIjQvuC7gyE6hol0; NID_JKL=9AmEcwfMQE0cYzXMkcf5fLLiUd+yxWR3svcyEyt8moY=; CBI_SNS=naver|mSZUwX9yJa5EY6hd; CBI_SES=X5tpj5h5j/50OmSRq3rC303XG/YlPaiAq9Dh0iB+8aEVtNwNfB2fEwbMp36LJCIsAjFiXqBR+ZcR6Fm0MqPInAuRN9j/k38Q+jP1MpGMz1JnZe8dbEGVSRvpBT6vXbJZe40tdCyd8tAWFz/shvS8xRV5p3kTWrGlpfaEmNnZEFfBbQq8D3V6wfpR1SChyerRt98IGH8qslI7tegndfK10f/+WgMiFqZxIsWj0XgGVyxho0tM8Q9vcAvvLgfx+dyudvnco1QDxE+ixJWJmsG1bP9g2XsQyEY+WxJoqBspFByXhqmU5PxEPTjP7CEyBCNkOSQK/AbbBzZ5XZJZWlDTcCaeP/ykmnKWk/kvDaZTmAH3adwO0Vb7cjyD6obOB2URByMgAF6tZ2Suep3J3jbxycknJk/zotAbjQ7bsKwDwXzuH7SzWE8yE/uBXiSDjuB2sRxFuEUjruj3jVNGlW79GA==; CBI_CHK="r5V0mf9uRUZHZ/vmLGy3ez7f4/k4aqWXL5o03eN68foZy3oPS8xU0v9iEv36wMxiHNPFJKIqnS8jaj+67l0NWg4BY2+gOeAANxCkpUdZggtLPOganbD/HumMEE/+hyh5kYpO8HQWLy3pTXrffuPRqqh9YGJAbleDrO2flGDOKc0="; NID_SES=AAABoBssrRTBsd6G0Ycqw/VBy0VWP1Yzfezk1ZK1rQk/7tt2R1YXYdgUwAt8guRUK+J/zToY6oJXZHLhBgStyB5K/K7uMULNCmbq50a1k7XuF1wtQNR3BGMcsGmMQlldFi0ORGahWPt1YJALKgtsqD9E0I6l7xU7P+BY5qToQEyl1YN0HZvxT3N0GL2pQV2mvrwiTetzeu32LEhSoUcqt1lrV16ZLN3lK8A4cileOfwfy9bdheFGSU8dstOA/7ndI0MPheCSe5MzTn1zR0/O0ApCJtqdETlylwJq68MvmghVfI/sL+vjQkFJScPKaxhCtnTLi8vFnbpixQg8hzISLHuUOlkfXwFeiNuPA6aftIzPKW3NDSWZhu8cPVHv2K8mBod9wjEB0XDeVgoSolCUJSGL60McAhhHzK983LxP4VqUhLJoyBep1CAHFXLK3WhjhKdvlATL72w88baVPSkMbCWUeNrt5wvPTtNkcoYGFpjEVAgAliJtzyjwG5MK5mSfvbYuu61KjZOHIOHlNeHrL2pUEBqwLFFERXZ4/KFK9eskkjgX; page_uid=i4UdIwprvTossB36spsssssstdV-461708; JSESSIONID=E7D53BE263AEF89A6E3A0FA122AA0C6E.jvm1',
            'referer': 'https://section.blog.naver.com/Search/Post.naver?pageNo=1&rangeType=MONTH&orderBy=sim&startDate=2023-05-11&endDate=2023-06-11&keyword=%EC%8A%A4%ED%8E%98%EC%9D%B8%20%EC%97%AC%ED%96%89%EC%A4%80%EB%B9%84',
            'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': '"Windows"',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-origin',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
        }

        params = {
            'countPerPage': '20',
            'currentPage': str(count),
            'endDate': '2023-06-11',
            'keyword': str(keyword),
            'orderBy': 'sim',
            'startDate': '2023-05-11',
            'type': 'post',
        }

        response = requests.get('https://section.blog.naver.com/ajax/SearchList.naver', params=params, cookies=cookies, headers=headers)
        # print(response.text)
        positionFr=response.text.find(',')
        resultRaw=response.text[positionFr+1:]
        # print(resultRaw)
        results=json.loads(resultRaw)['result']['searchList']
        if len(results)==0:
            print("데이타없음")
            break
        for result in results:
            print('데이타갯수:',len(dataList))
            title=result['noTagTitle']
            postUrl=result['postUrl']
            positionRR=postUrl.rfind("/")
            postUrl2=postUrl[:positionRR].replace("//blog","//m.blog")
            nickname=result['nickName']
            positionMail=postUrl2.rfind("/")
            postId=postUrl2[positionMail+1:]
            mailAddress=postUrl2[positionMail+1:]+"@naver.com"
            print(title,"/",nickname,"/",postUrl2)
            data={'keyword':keyword,'nickName':nickname,'postUrl':postUrl2,'mailAddress':mailAddress,'postId':postId}
            dataListPrev=dataList.copy()
            dataList.append(data)
            print(dataList)
            print(data)
            if data in dataListPrev:
                print("포함됨")
                sameCount=sameCount+1
            else:
                print("포함안됨")
                sameCount=0
            if sameCount>=5:
                endFlag=True
                break
            dataPrev=data
        if endFlag==True:
            break
        count=count+1
        time.sleep(random.randint(5,10)*0.1)
    return dataList

def getDetail(dataList):
    cookies = {
        'NNB': 'DSFIMBNTPBPGI',
        'nx_ssl': '2',
        'ASID': 'd3d7bf49000001881cf2740500000046',
        'nid_inf': '-1284035119',
        'NID_AUT': 'WuLSEC5nK+ib0TSryTGY6GlKj08nYB+0Nt1L4EutHlgWTzp+zIjQvuC7gyE6hol0',
        'NID_JKL': '9AmEcwfMQE0cYzXMkcf5fLLiUd+yxWR3svcyEyt8moY=',
        'CBI_SNS': 'naver|mSZUwX9yJa5EY6hd',
        'CBI_SES': 'X5tpj5h5j/50OmSRq3rC303XG/YlPaiAq9Dh0iB+8aEVtNwNfB2fEwbMp36LJCIsAjFiXqBR+ZcR6Fm0MqPInAuRN9j/k38Q+jP1MpGMz1JnZe8dbEGVSRvpBT6vXbJZe40tdCyd8tAWFz/shvS8xRV5p3kTWrGlpfaEmNnZEFfBbQq8D3V6wfpR1SChyerRt98IGH8qslI7tegndfK10f/+WgMiFqZxIsWj0XgGVyxho0tM8Q9vcAvvLgfx+dyudvnco1QDxE+ixJWJmsG1bP9g2XsQyEY+WxJoqBspFByXhqmU5PxEPTjP7CEyBCNkOSQK/AbbBzZ5XZJZWlDTcCaeP/ykmnKWk/kvDaZTmAH3adwO0Vb7cjyD6obOB2URByMgAF6tZ2Suep3J3jbxycknJk/zotAbjQ7bsKwDwXzuH7SzWE8yE/uBXiSDjuB2sRxFuEUjruj3jVNGlW79GA==',
        'NID_SES': 'AAABoBssrRTBsd6G0Ycqw/VBy0VWP1Yzfezk1ZK1rQk/7tt2R1YXYdgUwAt8guRUK+J/zToY6oJXZHLhBgStyB5K/K7uMULNCmbq50a1k7XuF1wtQNR3BGMcsGmMQlldFi0ORGahWPt1YJALKgtsqD9E0I6l7xU7P+BY5qToQEyl1YN0HZvxT3N0GL2pQV2mvrwiTetzeu32LEhSoUcqt1lrV16ZLN3lK8A4cileOfwfy9bdheFGSU8dstOA/7ndI0MPheCSe5MzTn1zR0/O0ApCJtqdETlylwJq68MvmghVfI/sL+vjQkFJScPKaxhCtnTLi8vFnbpixQg8hzISLHuUOlkfXwFeiNuPA6aftIzPKW3NDSWZhu8cPVHv2K8mBod9wjEB0XDeVgoSolCUJSGL60McAhhHzK983LxP4VqUhLJoyBep1CAHFXLK3WhjhKdvlATL72w88baVPSkMbCWUeNrt5wvPTtNkcoYGFpjEVAgAliJtzyjwG5MK5mSfvbYuu61KjZOHIOHlNeHrL2pUEBqwLFFERXZ4/KFK9eskkjgX',
        'page_uid': 'i4UdIwprvTossB36spsssssstdV-461708',
        'stat_yn': '1',
        'JSESSIONID': 'D1B56B6CBB4CC0995CFC523CAB828749.jvm1',
        'CBI_CHK': '"r5V0mf9uRUZHZ/vmLGy3ez7f4/k4aqWXL5o03eN68foZy3oPS8xU0v9iEv36wMxiHNPFJKIqnS8jaj+67l0NWg4BY2+gOeAANxCkpUdZggtLPOganbD/HumMEE/+hyh5ZsSiRkNHbgc+A/yAw1OChdi+XwPeMqIMjaCJrRjd7J4="',
    }

    headers = {
        'authority': 'm.blog.naver.com',
        'accept': 'application/json, text/plain, */*',
        'accept-language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7',
        # 'cookie': 'NNB=DSFIMBNTPBPGI; nx_ssl=2; ASID=d3d7bf49000001881cf2740500000046; nid_inf=-1284035119; NID_AUT=WuLSEC5nK+ib0TSryTGY6GlKj08nYB+0Nt1L4EutHlgWTzp+zIjQvuC7gyE6hol0; NID_JKL=9AmEcwfMQE0cYzXMkcf5fLLiUd+yxWR3svcyEyt8moY=; CBI_SNS=naver|mSZUwX9yJa5EY6hd; CBI_SES=X5tpj5h5j/50OmSRq3rC303XG/YlPaiAq9Dh0iB+8aEVtNwNfB2fEwbMp36LJCIsAjFiXqBR+ZcR6Fm0MqPInAuRN9j/k38Q+jP1MpGMz1JnZe8dbEGVSRvpBT6vXbJZe40tdCyd8tAWFz/shvS8xRV5p3kTWrGlpfaEmNnZEFfBbQq8D3V6wfpR1SChyerRt98IGH8qslI7tegndfK10f/+WgMiFqZxIsWj0XgGVyxho0tM8Q9vcAvvLgfx+dyudvnco1QDxE+ixJWJmsG1bP9g2XsQyEY+WxJoqBspFByXhqmU5PxEPTjP7CEyBCNkOSQK/AbbBzZ5XZJZWlDTcCaeP/ykmnKWk/kvDaZTmAH3adwO0Vb7cjyD6obOB2URByMgAF6tZ2Suep3J3jbxycknJk/zotAbjQ7bsKwDwXzuH7SzWE8yE/uBXiSDjuB2sRxFuEUjruj3jVNGlW79GA==; NID_SES=AAABoBssrRTBsd6G0Ycqw/VBy0VWP1Yzfezk1ZK1rQk/7tt2R1YXYdgUwAt8guRUK+J/zToY6oJXZHLhBgStyB5K/K7uMULNCmbq50a1k7XuF1wtQNR3BGMcsGmMQlldFi0ORGahWPt1YJALKgtsqD9E0I6l7xU7P+BY5qToQEyl1YN0HZvxT3N0GL2pQV2mvrwiTetzeu32LEhSoUcqt1lrV16ZLN3lK8A4cileOfwfy9bdheFGSU8dstOA/7ndI0MPheCSe5MzTn1zR0/O0ApCJtqdETlylwJq68MvmghVfI/sL+vjQkFJScPKaxhCtnTLi8vFnbpixQg8hzISLHuUOlkfXwFeiNuPA6aftIzPKW3NDSWZhu8cPVHv2K8mBod9wjEB0XDeVgoSolCUJSGL60McAhhHzK983LxP4VqUhLJoyBep1CAHFXLK3WhjhKdvlATL72w88baVPSkMbCWUeNrt5wvPTtNkcoYGFpjEVAgAliJtzyjwG5MK5mSfvbYuu61KjZOHIOHlNeHrL2pUEBqwLFFERXZ4/KFK9eskkjgX; page_uid=i4UdIwprvTossB36spsssssstdV-461708; stat_yn=1; JSESSIONID=D1B56B6CBB4CC0995CFC523CAB828749.jvm1; CBI_CHK="r5V0mf9uRUZHZ/vmLGy3ez7f4/k4aqWXL5o03eN68foZy3oPS8xU0v9iEv36wMxiHNPFJKIqnS8jaj+67l0NWg4BY2+gOeAANxCkpUdZggtLPOganbD/HumMEE/+hyh5ZsSiRkNHbgc+A/yAw1OChdi+XwPeMqIMjaCJrRjd7J4="',
        'referer': 'https://m.blog.naver.com/flfkfma',
        'sec-ch-ua': '"Not.A/Brand";v="8", "Chromium";v="114", "Google Chrome";v="114"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
        'sec-fetch-dest': 'empty',
        'sec-fetch-mode': 'cors',
        'sec-fetch-site': 'same-origin',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36',
    }

    response = requests.get('https://m.blog.naver.com/api/blogs/{}'.format(dataList['postId']), cookies=cookies, headers=headers)
    # print(response.text)
    result=json.loads(response.text)
    # pprint.pprint(result)
    dayVisitorCount=result['result']['dayVisitorCount']
    return dayVisitorCount




class Thread(QThread):
    cnt = 0
    user_signal = pyqtSignal(str)  # 사용자 정의 시그널 2 생성

    def __init__(self, parent,keyword):  # parent는 WndowClass에서 전달하는 self이다.(WidnowClass의 인스턴스)
        super().__init__(parent)
        self.parent = parent  # self.parent를 사용하여 WindowClass 위젯을 제어할 수 있다.
        self.keyword=keyword

    def run(self):
        keyword = self.keyword
        keywordList=keyword.split(",")
        for keywordElem in keywordList:
            text = "키워드:{}, 블로그 URL 탐색중...".format(keywordElem)
            print(text)
            self.user_signal.emit(text)
            dataList = getUrl(keywordElem)
            newDataList = []
            for dataElem in dataList:
                if dataElem not in newDataList:
                    newDataList.append(dataElem)
            dataList = newDataList

            text = "방문자수 조회중..."
            print(text)
            self.user_signal.emit(text)
            for dataElem in dataList:
                try:
                    dayVisitorCount = getDetail(dataElem)
                    dataElem.update({'dayVisitorCount': dayVisitorCount})
                except:
                    dataElem.update({'dayVisitorCount': 0})
                    time.sleep(10)
                print(dataElem)
                time.sleep(random.randint(5, 10) * 0.1)

            text = "엑셀 저장중..."
            print(text)
            self.user_signal.emit(text)
            wb = openpyxl.Workbook()
            ws = wb.active
            columnName = ['키워드', '닉네임', '블로그주소', '네이버메일주소', 'TODAY']
            ws.append(columnName)
            for dataElem in dataList:
                data = [dataElem['keyword'], dataElem['nickName'], dataElem['postUrl'], dataElem['mailAddress'],
                        dataElem['dayVisitorCount']]
                ws.append(data)
            timeNow = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            wb.save('result_{}_{}.xlsx'.format(keywordElem, timeNow))
            text = "작업 완료"
            print(text)
            self.user_signal.emit(text)

    def stop(self):
        pass

class Example(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.path = "C:"
        self.index = None
        self.setupUi(self)
        self.setSlot()
        self.show()
        QApplication.processEvents()

    def start(self):
        self.keyword=self.lineEdit.text()
        print("키워드는:",self.keyword)
        self.x = Thread(self,self.keyword)
        self.x.user_signal.connect(self.slot1)  # 사용자 정의 시그널2 슬롯 Connect
        self.x.start()

    def slot1(self, data1):  # 사용자 정의 시그널1에 connect된 function
        self.textEdit.append(str(data1))

    def setSlot(self):
        pass

    def setIndex(self, index):
        pass

    def quit(self):
        QCoreApplication.instance().quit()


app = QApplication([])
ex = Example()
sys.exit(app.exec_())





