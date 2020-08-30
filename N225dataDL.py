#!/usr/bin/env python
# coding: utf-8
#JPXのサイトから特定のファイルをダウンロードし加工するプログラムを作成する
#デリバティブ建玉残高表は毎営業日更新されるため、直近のデータ以外は収集できない。データ加工を自動化できるとDLし忘れがなく効率化が可能
#１．JPXサイトにアクセスし特定のファイルをダウンロードする
#２．ダウンロードディレクトリから任意のディレクトリへファイルを移動する
#3.　デリバティブ建玉残高表から必要なデータを別のエクセルに張り付ける



#１．以下、JPXのサイトに接続し、デリバティブ建玉残高表をダウンロードするためのコード
from selenium import webdriver
import time
#chrome自動操作のためのwebdriberのパスを設定
#webdriberの取得と設定に関してはキノコードさんのyoutubeを参照
page01 = webdriver.Chrome(executable_path = 'C:\\Users\\sato\\Desktop\\mypandas\\chromedriver.exe')
#待機3秒
page01.implicitly_wait(3)
# JPXの当日取引高等のページにアクセス
url01 = "https://www.jpx.co.jp/markets/derivatives/trading-volume/index.html"
page01.get(url01)
#待機3秒
time.sleep(3)
#OSEデリバティブ建玉残高表のダウンロード　
#HTMLのボタンにIDやname属性がないためXパスでダウンロードボタンを指定
#Xパスの取得についてはキノコードさんのyoutubeを参照
dataDL = page01.find_element_by_xpath('/html/body/div[2]/div[4]/div[1]/div[3]/div[3]/div/table/tbody/tr[2]/td[3]/a')
time.sleep(1)
dataDL.click()
#ダウンロード時間に余裕を持たせるための待機
time.sleep(10)
#以上、ソースになるファイルのダウンロード完了



#２．以下、ダウンロードファイルを任意のディレクトリへ移動するためのコード
#OSEデリバティブ建玉残高表のファイル名に日付が含まれているため、当日の日付をファイル名に付けるように加工する
import datetime
day = datetime.date.today()
#本日の日付であるdatetime型のデータ(2020-08-28)からハイフンを抜く
dayName = day.strftime('%Y%m%d')
#日付＋共通ファイル名
#OSEデリバティブ建玉残高表は毎営業日の20:00に更新されるため、20:00以前および土日祝日にプログラムを実行するとここでエラーになるので注意
fileName = dayName+'open_interest.xls'
#ファイル移動のためのインポート
import shutil
#ダウンロードディレクトリから指定ディレクトリへ移動　
#ここではDownloads⇒datasource　設定に応じて任意のディレクトリへ変更
shutil.move('C:\\Users\\sato\\Downloads\\'+fileName, 'C:\\Users\\sato\\Desktop\\datasource')
#以上、ソースになるダウンロードファイルの移動完了



#３.以下、デリバティブ建玉残高表から必要なデータをvolume&open_contract.xlsx張り付ける
import pandas as pd
import openpyxl as px
#デリバティブ建玉残高表の拡張子が.xlsであるため、excel2003以前のエクセルを操作するためのxlrdライブラリをインポート
import xlrd
#デリバティブ建玉残高表の読み込み
ds = xlrd.open_workbook('C:\\Users\\sato\\Desktop\\datasource\\'+fileName)
#シートの読み込み
sheet01 = ds.sheet_by_name('デリバティブ建玉残高状況')
#取引高のコピー
n225Vol = sheet01.cell(36,2)    #N225ラージ
n225minVol = sheet01.cell(33,9) #N225miniの取引高
topixVol = sheet01.cell(42,2)   #topixの取引高
putVol = sheet01.cell(147,2)    #プットオプションの取引高
callVol = sheet01.cell(148,2)   #コールオプションの取引高
