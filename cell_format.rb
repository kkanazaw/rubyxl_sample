# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルに文字を追加
cell = worksheet.add_cell(0, 1, 12345)
# 書式設定
# see https://support.office.com/en-us/article/5026bbd6-04bc-48cd-bf33-80f18b4eae68
cell.set_number_format('#,##0')
# sample.xlsxという名前で保存
workbook.write('outputs/cell_format.xlsx')
