#!/usr/bin/env python
# coding: utf-8

#以下、JPXのサイトに接続し、デリバティブ建玉残高表をダウンロード
from selenium import webdriver
import time
import pandas as pd
#chrome自動操作のためのwebdriberのパス
page01 = webdriver.Chrome(executable_path = 'C:\\Users\\sato\\Desktop\\mypandas\\chromedriver.exe')
#待機3秒
page01.implicitly_wait(3)
# JPXの当日取引高等のページにアクセス
url01 = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
page01.get(url01)
#待機3秒
time.sleep(3)
#OSEデリバティブ建玉残高表のダウンロード　
#IDやname属性がないためxパスでダウンロードボタンを指定
dataDL = page01.find_element_by_xpath('/html/body/div[2]/div[4]/div[1]/div[3]/div[3]/div/table/tbody/tr[2]/td[3]/a')
time.sleep(1)
dataDL.click()
time.sleep(10)
#以上、ファイルのダウンロード完了

#以下、ダウンロードファイルをダウンロードディレクトリから特定のディレクトリへ移動
#ファイル名に日付が含まれているため、当日の日付をファイル名に付けるように加工
import datetime
day = datetime.date.today()
dayName = day.strftime('%Y%m%d')
#日付＋共通ファイル名
fileName = dayName+'open_interest.xls'
#ファイル移動のためのインポート
import shutil
#ダウンロードディレクトリから指定ディレクトリへ移動
shutil.move('C:\\Users\\sato\\Downloads\\'+fileName, 'C:\\Users\\sato\\Desktop\\datasource')
