# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# シートの名前を変更
worksheet.sheet_name = '新しいシート名'
# セルに文字を追加
worksheet.add_cell(0, 1, 'B1セル')
# 2シート目を追加
worksheet = workbook.add_worksheet('次のシート')
# sample.xlsxという名前で保存
workbook.write('outputs/add_sheet.xlsx')
