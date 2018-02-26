import openpyxl

# 新規スプレットシートとして取得
workbook = openpyxl.Workbook()
sheet    = workbook.active

# サンプルのエクセル通りにデータを作る
sheet['A1'] = '2015/4/5 13:34'
sheet['A2'] = '2015/4/5 3:41'
sheet['A3'] = '2015/4/6 12:46'
sheet['A4'] = '2015/4/8 8:59'
sheet['A5'] = '2015/4/10 2:07'
sheet['A6'] = '2015/4/10 18:10'
sheet['A7'] = '2015/4/10 2:40'
sheet['B1'] = 'Apples'
sheet['B2'] = 'Cherries'
sheet['B3'] = 'Pears'
sheet['B4'] = 'Oranges'
sheet['B5'] = 'Apples'
sheet['B6'] = 'Bananas'
sheet['B7'] = 'Strawberries'
sheet['C1'] = 73
sheet['C2'] = 85
sheet['C3'] = 14
sheet['C4'] = 52
sheet['C5'] = 152
sheet['C6'] = 23
sheet['C7'] = 98

# 最後の行にC列のSUM情報を追加する
sheet['C8'] = '=SUM(C1:C7)'

# 保存
workbook.save('./excel/python_write.xlsx')