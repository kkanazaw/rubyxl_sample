# coding:utf-8
require 'rubyXL'
# 新しいworkbookの作成
workbook = RubyXL::Workbook.new
# worksheetを取得
# デフォルトで1つ生成されている
worksheet = workbook[0]
# セルにチェックマークが追加
# wikipedia https://ja.wikipedia.org/wiki/%E3%83%81%E3%82%A7%E3%83%83%E3%82%AF%E3%83%9E%E3%83%BC%E3%82%AF
cell = worksheet.add_cell(1, 0, '☑') #チェックマーク U+2611 BALLOT BOX WITH CHECK
cell2 = worksheet.add_cell(2, 0, '□') # 四角
cell3 = worksheet.add_cell(3, 0, '🗹')  # U+1F5F9 BALLOT BOX WITH BOLD CHECK
# sample.xlsxという名前で保存
workbook.write('outputs/checkbox.xlsx')
