# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルに文字を追加
cell = worksheet.add_cell(0, 1, '')
# セルの下に罫線を引く
cell.change_border(:bottom, 'medium')
# 細い罫線を引く
# チェインメソッドな指定も可能
cell = worksheet.add_cell(0, 3, '')
       .change_border(:bottom, 'thin')
# sample.xlsxという名前で保存
workbook.write('outputs/border.xlsx')
