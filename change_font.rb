# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルに文字を追加
cell = worksheet.add_cell(0, 1, 'B1')
# フォントを変更
cell.change_font_name '梅明朝'
# フォントサイズを変更
cell.change_font_size 20
# フォントの色を変更
cell.change_font_color 'ff0000'
# セルの色を変更
cell.change_fill '00ff00'
# sample.xlsxという名前で保存
workbook.write('outputs/change_font.xlsx')
