# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルに文字を追加
worksheet.merge_cells(0, 0, 1, 1)
# 幅を広げる
# sample.xlsxという名前で保存
workbook.write('outputs/merge_cells.xlsx')
