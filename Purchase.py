import pandas as pd
#sekkdfに必要なカラム番号だけのデータフレームを売上.xlsから抽出する
selldf = pd.read_excel('C:\\Users\\Himuka\\Desktop\\purchaseData\\売上.xls',usecols=[1,5,6,9,11,12,15,17,18])

#procuctRistの作成　商品一覧　重複を削除し、ユニークデーターを抽出
productList = selldf[['取引区分','商品名']].drop_duplicates()
#purchaseRistの作成　仕入れ一覧　取引区分の「掛その他」で抽出
purchaseList = productList[productList['取引区分'] == '掛その他']
#仕入れ一覧の中からトマト、レタスを区分
tomatoPurchaseList = purchaseList[purchaseList['商品名'].str.contains('トマト|日向|ｷｬﾛﾙ|桃姫|ﾄﾏﾄ')]
lettucePurchaseList = purchaseList[purchaseList['商品名'].str.contains('レタス|ﾚﾀｽ')]

#返品以外のトマト商品のリスト
x = tomatoPurchaseList [~tomatoPurchaseList ['商品名'].str.contains('返品')]
simpleListT = x['商品名']
#返品以外のレタス仕入商品のリスト
x = lettucePurchaseList [~lettucePurchaseList ['商品名'].str.contains('返品')]
simpleListL = x['商品名']



#トマト仕入れ商品の売上データを商品名、得意先、売上単価ごとにグループ化する(返品以外)
tomatoSellData = []
for i in simpleListT:
    nam = selldf[selldf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['得意先ｺｰﾄﾞ','得意先名','取引年月日','商品名','売上単価']).sum()
    tomatoSellData.append(nam)

#レタスの売上データを商品名、得意先、売上単価ごとにグループ分けする(返品以外)
lettuceSellData = []
for i in simpleListL:
    nam = selldf[selldf['商品名']== i]
    nam = nam.groupby(['得意先ｺｰﾄﾞ','得意先名','取引年月日','商品名','売上単価']).sum()
    lettuceSellData.append(nam)



#仕入れエクセルから必要な列のみを読み込む
buydf = pd.read_excel('C:\\Users\\Himuka\\Desktop\\purchaseData\\仕入れ.xls',usecols=[1,5,6,9,11,12,15,17,18,21,22])

#仕入一覧からトマトのデータを抽出し、仕入先・商品・仕入単価ごとにグループ化する
tomatoBuyData = []
for i in simpleListT:
    nam = buydf[buydf['商品名']== i]
    nam = nam.groupby(['仕入先ｺｰﾄﾞ','仕入先名','取引年月日','商品名','仕入単価']).sum()
    tomatoBuyData.append(nam)

#仕入れ一覧から仕入レタスのデータを抽出し、仕入先・商品・仕入単価ごとにグループ化する
lettuceBuyData = []
for i in simpleListL:
    nam = buydf[buydf['商品名']== i]
    nam = nam.groupby(['仕入先ｺｰﾄﾞ','仕入先名','取引年月日','商品名','仕入単価','行摘要']).sum()
    lettuceBuyData.append(nam)



import openpyxl
import datetime
filePath = 'C:\\Users\\Himuka\\Desktop\\sellData\\仕入商品売上データ.xlsx'
sellData = openpyxl.load_workbook(filePath)



tomatoSheet = sellData['tomato']
i = 3
for nam in simpleListT:
    tomatoSheet.cell(row=i,column=3).value = nam
    i += 1



lettuceSheet = sellData['lettuce']
i = 3
for nam in simpleListL:
    lettuceSheet.cell(row=i,column=3).value = nam
    i += 1


sellData.save(filePath)