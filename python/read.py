import openpyxl

# ファイル読み込み
workbook = openpyxl.load_workbook('./excel/example.xlsx')

# データを表示
for row in workbook.active:
    print(row[0].value, row[1].value, row[2].value, sep="\t")

