import pandas as pd
#�f�[�^�t�@�C���ǂݍ��ݗp�̃t�H���_�p�X
inputFolderPath = 'C:\\Users\\Himuka\\Desktop\\purchaseData\\'
#�f�[�^�o�͗p�̃t�H���_�p�X
outputFolderPath = 'C:\\Users\\Himuka\\Desktop\\sellData\\'
################����DB�쐬#####################################
#selldf�ɕK�v�ȃJ�����ԍ������̃f�[�^�t���[���𔄏�.xls���璊�o����
selldf = pd.read_excel(inputFolderPath+'����.xls',usecols=[1,5,6,9,11,12,15,17,18])
#�����l���f�Ȃ��f�Œu��������
selldf = selldf.fillna('�Ȃ�')

#procuctRist�̍쐬�@���i�ꗗ�@�d�����폜���A���j�[�N�f�[�^�[�𒊏o
sellProductList = selldf[['����敪','���i��']].drop_duplicates()
#purchaseRist�̍쐬�@�d����ꗗ�@����敪�́u�|���̑��v�Œ��o
sellPurchaseList = sellProductList[sellProductList['����敪'] == '�|���̑�']
#�d���ꏤ�i�ꗗ�̒�����g�}�g�A���^�X���敪
tom = sellPurchaseList[sellPurchaseList['���i��'].str.contains('�g�}�g|����|����|���P|���')]
let = sellPurchaseList[sellPurchaseList['���i��'].str.contains('���^�X|���')]


#�ԕi�ȊO�̃g�}�g���i�̃��X�g
x = tom [~tom ['���i��'].str.contains('�ԕi')]
sell_uniqueTomato = x['���i��']
#�ԕi�ȊO�̃��^�X�d�����i�̃��X�g
x = let [~let ['���i��'].str.contains('�ԕi')]
sell_uniqueLettuce = x['���i��']

#�ԕi���X�g
x = tom[tom ['���i��'].str.contains('�ԕi')]
sell_returne_uniqueTomato = x['���i��']
#�ԕi�ȊO�̃��^�X�d�����i�̃��X�g
x = let [let ['���i��'].str.contains('�ԕi')]
sell_returne_uniqueLettuce = x['���i��']

#�g�}�g�d���ꏤ�i�̔���f�[�^�����i���A���Ӑ�A����P�����ƂɃO���[�v������(�ԕi�ȊO)
selldf_tomato_array = []
for i in sell_uniqueTomato:
    nam = selldf[selldf['���i��']== i]
    selldf_tomato_array.append(nam)

#���^�X�̔���f�[�^�����i���A���Ӑ�A����P�����ƂɃO���[�v��������(�ԕi�ȊO)
selldf_lettuce_array = []
for i in sell_uniqueLettuce:
    nam = selldf[selldf['���i��']== i]
    selldf_lettuce_array.append(nam)

#�g�}�g����ԕi���X�g�̍쐬
selldf_returneTomato_array =[]
for i in sell_returne_uniqueTomato:
    nam = selldf[selldf['���i��']== i]
    selldf_returneTomato_array.append(nam)

#�d���ꃌ�^�X����ԕi���X�g�̍쐬
selldf_returneLettuce_array =[]
for i in sell_returne_uniqueLettuce:
    nam = selldf[selldf['���i��']== i]
    selldf_returneLettuce_array.append(nam)