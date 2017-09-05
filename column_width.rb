# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルに文字を追加
cell = worksheet.add_cell(0, 1, '長い文字列ですよー')
# 幅を広げる
worksheet.change_column_width(1, 20)
# sample.xlsxという名前で保存
workbook.write('outputs/column_width.xlsx')
