# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# フォントのデフォルト指定
workbook.fonts[0].set_name('梅明朝')
workbook.fonts[0].set_size(14)
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルに文字を追加
cell = worksheet.add_cell(0, 1, '日本語')
# sample.xlsxという名前で保存
workbook.write('outputs/default_format.xlsx')
