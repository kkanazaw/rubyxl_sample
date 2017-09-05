# coding: utf-8
require 'rubyXL'

# ワークブックの読込
workbook = RubyXL::Parser.parse('excels/beaglesoft_template2.xlsx')

# コピー元シート
org_sheet =  workbook['テストシート名']

add_sheet = Marshal.load(Marshal.dump(org_sheet))

# こちらの方法でも問題ないが、エラーとなるケースが存在したためコメントアウト
# Marshal.dump でシリアライズできない情報が欠落している可能性もある。
# add_sheet = org_sheet

add_sheet.sheet_name = 'コピーシート名'
add_sheet.workbook = workbook

# worksheetの番号を取得する
add_sheet.sheet_id = workbook.worksheets.map(&:sheet_id).max + 1

# テストシートをコピーしてブックに追加
workbook.worksheets << add_sheet

# ブックの内容を列挙する
workbook.worksheets.each do |worksheet|
  p "シートのクラス名：#{worksheet.class.name}"
  p "シート名：#{worksheet.sheet_name}"
end

# 元から存在するシートを削除する
delete_sheet_index = workbook.worksheets.index(org_sheet)
workbook.worksheets.delete_at delete_sheet_index

# outputsディレクトリに保存する
file_name = "beaglesoft_template_#{Time.now.strftime('%Y%m%d%H%M%s')}.xlsx"
workbook.write("outputs/#{file_name}")
