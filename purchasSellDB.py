#伝票作成システムからエクスポートしたデータをもとに仕入商品ごとの売り上げデータベースを作成する

import pandas as pd
folderPath = ''C:\\Users\\sato\\Desktop\\purchaseData\\'
################売上DB作成#####################################
#selldfに必要なカラム番号だけのデータフレームを売上.xlsから抽出する
selldf = pd.read_excel(folderPath+'売上.xls',usecols=[1,5,6,9,11,12,15,17,18])
#欠損値を’なし’で置き換える
selldf = selldf.fillna('なし')

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

#返品リスト
x = tomatoPurchaseList [tomatoPurchaseList ['商品名'].str.contains('返品')]
returnedT = x['商品名']
#返品以外のレタス仕入商品のリスト
x = lettucePurchaseList [lettucePurchaseList ['商品名'].str.contains('返品')]
returnedL = x['商品名']



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



#売上データDB作成
import openpyxl
import datetime
day = datetime.date.today()
dayName = day.strftime('%Y%m%d')

#トマト販売データベース
fileName = dayName + 'トマト販売データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in simpleListT:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in simpleListT:
        tomatoSellData[j].to_excel(writer, sheet_name=i)
        j += 1

#売上列の3桁カンマ区切り
wb = openpyxl.load_workbook(filePath)
for i in simpleListT:
    sh = wb[i]
    a = sh.max_row
    for b in range(a):
        sh.cell(row=b+1,column=6).number_format = '#,##0'
wb.save(filePath)



#レタス販売データベース
fileName = dayName + '仕入レタス販売データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in simpleListL:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in simpleListL:
        lettuceSellData[j].to_excel(writer, sheet_name=i)
        j += 1
#売上列の3桁カンマ区切り
wb = openpyxl.load_workbook(filePath)
for i in simpleListL:
    sh = wb[i]
    a = sh.max_row
    for b in range(a):
        sh.cell(row=b+1,column=6).number_format = '#,##0'
wb.save(filePath)


#トマト売上返品リストの作成
fileName = dayName + 'トマト売上返品データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in returnedT:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in returnedT:
        returnedSimpleListT[j].to_excel(writer, sheet_name=i)
        j += 1

#売上列の3桁カンマ区切り
wb = openpyxl.load_workbook(filePath)
for i in returnedT:
    sh = wb[i]
    a = sh.max_row
    for b in range(a):
        sh.cell(row=b+1,column=6).number_format = '#,##0'
wb.save(filePath)

#レタス売上返品データベース
fileName = dayName + '仕入レタス売上返品データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in returnedL:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in returnedL:
        returnedSimpleListL[j].to_excel(writer, sheet_name=i)
        j += 1

#売上列の3桁カンマ区切り
wb = openpyxl.load_workbook(filePath)
for i in returnedL:
    sh = wb[i]
    a = sh.max_row
    for b in range(a):
        sh.cell(row=b+1,column=6).number_format = '#,##0'
wb.save(filePath)
print('販売データ作成')

#######################売上DB作成以上####################



#######################仕入DB作成##########################
buydf = pd.read_excel(folderPath+'仕入れ.xls',usecols=[1,5,6,9,11,12,15,17,22], skip_blank_lines=False)
#欠損値を’なし’で置き換える
buydf = buydf.fillna('なし')

#仕入procuctRistの作成　商品一覧　重複を削除し、ユニークデーターを抽出
productListP = buydf[['取引区分','商品名']].drop_duplicates()
#仕入れ一覧の中からトマト、レタスを区分
PurchaseListT = productListP[productListP['商品名'].str.contains('トマト|日向|ｷｬﾛﾙ|桃姫|ﾄﾏﾄ')]
PurchaseListL = productListP[productListP['商品名'].str.contains('レタス|ﾚﾀｽ')]


#返品以外のトマト商品のリスト
x = PurchaseListT [~PurchaseListT ['商品名'].str.contains('返品')]
simpleListBT = x['商品名']
#返品以外のレタス仕入商品のリスト
x = PurchaseListL [~PurchaseListL ['商品名'].str.contains('返品')]
simpleListBL = x['商品名']

#返品リスト
x = PurchaseListT [PurchaseListT ['商品名'].str.contains('返品')]
returnedBT = x['商品名']
#返品以外のレタス仕入商品のリスト
x = PurchaseListL [PurchaseListL ['商品名'].str.contains('返品')]
returnedBL = x['商品名']


#トマト仕入れ商品の仕入データを商品名、得意先、売上単価ごとにグループ化する(返品以外)
tomatoBuyData = []
for i in simpleListBT:
    nam = buydf[buydf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['取引年月日','商品ｺｰﾄﾞ','仕入単価','仕入先ｺｰﾄﾞ','仕入先名','行摘要']).sum()
    tomatoBuyData.append(nam)

#仕入レタスの仕入れデータを商品名、得意先、売上単価ごとにグループ分けする(返品以外)
lettuceBuyData = []
for i in simpleListBL:
    nam = buydf[buydf['商品名']== i]
    nam = nam.groupby(['取引年月日','商品ｺｰﾄﾞ','仕入単価','仕入先ｺｰﾄﾞ','仕入先名','行摘要']).sum()
    lettuceBuyData.append(nam)

#トマト売上返品リストの作成
returnedSimpleListBT =[]
for i in returnedBT:
    nam = buydf[buydf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['取引年月日','商品ｺｰﾄﾞ','仕入単価','仕入先ｺｰﾄﾞ','仕入先名','行摘要']).sum()
    returnedSimpleListBT.append(nam)

#仕入れレタス売上返品リストの作成
returnedSimpleListBL =[]
for i in returnedBL:
    nam = buydf[buydf['商品名']== i]
    #得意先コード、得意先名、商品名、売上単価をキーにして、同じ売上・数量は合計する
    nam = nam.groupby(['取引年月日','商品ｺｰﾄﾞ','仕入単価','仕入先ｺｰﾄﾞ','仕入先名','行摘要']).sum()
    returnedSimpleListBL.append(nam)

#仕入DB作成
#トマト仕入データベース
fileName = dayName + 'トマト仕入データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in simpleListBT:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in simpleListBT:
        tomatoBuyData[j].to_excel(writer, sheet_name=i)
        j += 1



#レタス仕入データベース
fileName = dayName + '仕入レタス仕入データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in simpleListBL:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in simpleListBL:
        lettuceBuyData[j].to_excel(writer, sheet_name=i)
        j += 1


#トマト仕入返品リストの作成
fileName = dayName + 'トマト仕入返品データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in returnedBT:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in returnedBT:
        returnedSimpleListBT[j].to_excel(writer, sheet_name=i)
        j += 1



#レタス仕入返品データベース
fileName = dayName + '仕入レタス売上返品データ.xlsx'
filePath = folderPath+fileName
wb = openpyxl.Workbook()
for i in returnedBL:
    wb.create_sheet(i)
wb.save(filePath)
j = 0
with pd.ExcelWriter(filePath, engine='openpyxl') as writer:
    for i in returnedBL:
        returnedSimpleListBL[j].to_excel(writer, sheet_name=i)
        j += 1
print('仕入データを作成')
############################仕入Db作成終了####################
