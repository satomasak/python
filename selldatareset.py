import pandas as pd
#データファイル読み込み用のフォルダパス
inputFolderPath = 'C:\\Users\\Himuka\\Desktop\\purchaseData\\'
#データ出力用のフォルダパス
outputFolderPath = 'C:\\Users\\Himuka\\Desktop\\sellData\\'
################売上DB作成#####################################
#selldfに必要なカラム番号だけのデータフレームを売上.xlsから抽出する
selldf = pd.read_excel(inputFolderPath+'売上.xls',usecols=[1,5,6,9,11,12,15,17,18])
#欠損値を’なし’で置き換える
selldf = selldf.fillna('なし')

#procuctRistの作成　商品一覧　重複を削除し、ユニークデーターを抽出
sellProductList = selldf[['取引区分','商品名']].drop_duplicates()
#purchaseRistの作成　仕入れ一覧　取引区分の「掛その他」で抽出
sellPurchaseList = sellProductList[sellProductList['取引区分'] == '掛その他']
#仕入れ商品一覧の中からトマト、レタスを区分
tom = sellPurchaseList[sellPurchaseList['商品名'].str.contains('トマト|日向|ｷｬﾛﾙ|桃姫|ﾄﾏﾄ')]
let = sellPurchaseList[sellPurchaseList['商品名'].str.contains('レタス|ﾚﾀｽ')]


#返品以外のトマト商品のリスト
x = tom [~tom ['商品名'].str.contains('返品')]
sell_uniqueTomato = x['商品名']
#返品以外のレタス仕入商品のリスト
x = let [~let ['商品名'].str.contains('返品')]
sell_uniqueLettuce = x['商品名']

#返品リスト
x = tom[tom ['商品名'].str.contains('返品')]
sell_returne_uniqueTomato = x['商品名']
#返品以外のレタス仕入商品のリスト
x = let [let ['商品名'].str.contains('返品')]
sell_returne_uniqueLettuce = x['商品名']

#トマト仕入れ商品の売上データを商品名、得意先、売上単価ごとにグループ化する(返品以外)
selldf_tomato_array = []
for i in sell_uniqueTomato:
    nam = selldf[selldf['商品名']== i]
    selldf_tomato_array.append(nam)

#レタスの売上データを商品名、得意先、売上単価ごとにグループ分けする(返品以外)
selldf_lettuce_array = []
for i in sell_uniqueLettuce:
    nam = selldf[selldf['商品名']== i]
    selldf_lettuce_array.append(nam)

#トマト売上返品リストの作成
selldf_returneTomato_array =[]
for i in sell_returne_uniqueTomato:
    nam = selldf[selldf['商品名']== i]
    selldf_returneTomato_array.append(nam)

#仕入れレタス売上返品リストの作成
selldf_returneLettuce_array =[]
for i in sell_returne_uniqueLettuce:
    nam = selldf[selldf['商品名']== i]
    selldf_returneLettuce_array.append(nam)