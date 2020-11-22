#伝票作成システムからエクスポートしたデータをもとに仕入商品ごとの売り上げデータベースを作成する

import pandas as pd
#sekkdfに必要なカラム番号だけのデータフレームを売上.xlsから抽出する
selldf = pd.read_excel('C:\\Users\\sato\\Desktop\\purchaseData\\売上.xls',usecols=[1,5,6,9,11,12,15,17,18])

#商品基本データの抽出・加工
#procuctRistの作成　商品一覧　重複を削除し、ユニークデーターを抽出
productList = selldf[['取引区分','商品名']].drop_duplicates()
#purchaseRistの作成　仕入れ一覧　取引区分の「掛その他」で抽出
purchaseList = productList[productList['取引区分'] == '掛その他']
#売上一覧の中からトマト、レタスを区分
tomatoPurchaseList = purchaseList[purchaseList['商品名'].str.contains('トマト|日向|ｷｬﾛﾙ|桃姫|ﾄﾏﾄ')]
lettucePurchaseList = purchaseList[purchaseList['商品名'].str.contains('レタス|ﾚﾀｽ')]



#返品以外のトマト商品のリスト
x = tomatoPurchaseList [~tomatoPurchaseList ['商品名'].str.contains('返品')]
simpleListT = x['商品名']
#返品以外のレタス仕入商品のリスト
x = lettucePurchaseList [~lettucePurchaseList ['商品名'].str.contains('返品')]
simpleListL = x['商品名']

#返品リスト
x = tomatoPurchaseList [tomatoPurchaseList ['商品名'].str.contains('返品')]
returnedT = x['商品名']
#返品以外のレタス仕入商品のリスト
x = lettucePurchaseList [lettucePurchaseList ['商品名'].str.contains('返品')]
returnedL = x['商品名']


#売上データの加工　必要なものだけまとめる作業
#トマト仕入れ商品の売上データを商品名、得意先、売上単価ごとにグループ化する(返品以外)
tomatoSellData = []
for i in simpleListT:
    nam = selldf[selldf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['取引年月日','売上単価','得意先ｺｰﾄﾞ','得意先名']).sum()
    tomatoSellData.append(nam)

#レタスの売上データを商品名、得意先、売上単価ごとにグループ分けする(返品以外)
lettuceSellData = []
for i in simpleListL:
    nam = selldf[selldf['商品名']== i]
    nam = nam.groupby(['取引年月日','売上単価','得意先ｺｰﾄﾞ','得意先名']).sum()
    lettuceSellData.append(nam)

#トマト売上返品リストの作成
returnedSimpleListT =[]
for i in returnedT:
    nam = selldf[selldf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['取引年月日','売上単価','得意先ｺｰﾄﾞ','得意先名']).sum()
    returnedSimpleListT.append(nam)

#仕入れレタス売上返品リストの作成
returnedSimpleListL =[]
for i in returnedL:
    nam = selldf[selldf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['取引年月日','売上単価','得意先ｺｰﾄﾞ','得意先名']).sum()
    returnedSimpleListL.append(nam)



#売上データDB作成　パス先にデータベースになるエクセルを作成する
import openpyxl
import datetime
day = datetime.date.today()
dayName = day.strftime('%Y%m%d')

#トマト販売データベース
fileName = dayName + 'トマト販売データ.xlsx'
filePath = 'C:\\Users\\sato\\Desktop\\sellData\\'+fileName
wb = openpyxl.Workbook()
for i in simpleListT:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in simpleListT:
        tomatoSellData[j].to_excel(writer, sheet_name=i)
        j += 1

#レタス販売データベース
fileName = dayName + '仕入レタス販売データ.xlsx'
filePath = 'C:\\Users\\sato\\Desktop\\sellData\\'+fileName
wb = openpyxl.Workbook()
for i in simpleListL:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in simpleListL:
        lettuceSellData[j].to_excel(writer, sheet_name=i)
        j += 1

#トマト売上返品リストの作成
fileName = dayName + 'トマト売上返品データ.xlsx'
filePath = 'C:\\Users\\sato\\Desktop\\sellData\\'+fileName
wb = openpyxl.Workbook()
for i in returnedT:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in returnedT:
        returnedSimpleListT[j].to_excel(writer, sheet_name=i)
        j += 1

#レタス売上返品データベース
fileName = dayName + '仕入レタス売上返品データ.xlsx'
filePath = 'C:\\Users\\sato\\Desktop\\sellData\\'+fileName
wb = openpyxl.Workbook()
for i in returnedL:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in returnedL:
        returnedSimpleListL[j].to_excel(writer, sheet_name=i)
        j += 1
