# coding:utf-8
require 'rubyXL'
# あらかじめページにあわせるを設定したworkbookを読み込む
workbook = RubyXL::Parser.parse('inputs/fit_to_page.xlsx')
# worksheetを取得
worksheet = workbook[0]
# セルに文字を追加
cell = worksheet.add_cell(0, 1, 12345)
# sample.xlsxという名前で保存
workbook.write('outputs/fit_to_page.xlsx')
