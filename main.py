import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def count_strings(log_file):
    counts = {}
    with open(log_file, 'r') as file:
        for line in file:
            line = line.strip()
            if line:
                parts = line.split('”')
                if len(parts) >= 2:
                    target_string = parts[-2]
                    if target_string in counts:
                        counts[target_string] += 1
                    else:
                        counts[target_string] = 1
    return counts

log_file = 'log/productsearch-jp-202306_.log'
string_counts = count_strings(log_file)

# データをDataFrameに変換
df = pd.DataFrame.from_dict(string_counts, orient='index', columns=['検索回数'])

# Excelファイルに出力
output_file = 'excel/log_counts.xlsx'
writer = pd.ExcelWriter(output_file, engine='openpyxl')
df.to_excel(writer, sheet_name='Sheet1', index_label='検索文字列')

# シートのフォントと配置を変更
workbook = writer.book
worksheet = workbook['Sheet1']
font = Font(size=12, bold=False)  # サイズや太字の設定を変更する場合はここを調整
alignment = Alignment(horizontal='left')  # 左寄せに設定
for cell in worksheet['A']:
    cell.font = font
    cell.alignment = alignment

# 列の幅を調整
column_width = max([len(str(s)) for s in df.index]) + 3  # 文字列の長さに基づいて幅を設定
worksheet.column_dimensions['A'].width = column_width

writer.save()
print('Excelファイルに出力しました:', output_file)
