import openpyxl as xl

dataPath = 'C:\\Users\\Himuka\\Desktop\\sellData\\'
dbPath = 'C:\\Users\\Himuka\\Desktop\\�d���Ǘ�\\'
outPath = 'C:\\Users\\Himuka\\Desktop\\output\\'

otomemini = sellwb['�������Ƃ߃~�j']

#�\�̃^�C�g���쐬
wb = openpyxl.Workbook()
sheets = wb.create_sheet('�d�����i���v��')
title = ['���i��','����P��','���㐔','������z''�d���P��','�d������','�d�����z','�e��','�^��','�t�B-','�G��','���v���z','���v��']

j = 2
for i in title:
    sheets.cell(row=2, column=j).value = i
    j += 1
wb.save(outPath+'�d�����i���v��.xlsx')
outFile = outPath+'�d�����i���v��.xlsx'

