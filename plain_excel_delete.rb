# coding: utf-8
require 'rubyXL'

# ワークブックの読込
#template = RubyXL::Parser.parse('excels/プレーンExcelサンプル.xlsx')
workbook = RubyXL::Parser.parse('excels/プレーンExcelサンプル.xlsx')

# コピー元シート
del_sheet =  workbook['Sheet1']

# worksheetの番号を取得する
#add_sheet.sheet_id = workbook.worksheets.map(&:sheet_id).max + 1
#add_sheet.printer_settings[0].xlsx_path = ::Pathname.new("/xl/printerSettings/printerSettings#{add_sheet.sheet_id}.bin")

# テストシートをコピーしてブックに追加

#workbook.relationship_container.relationships[3].id = "rId7"
#workbook.relationship_container.relationships[4].id = "rId6"
#workbook.relationship_container.relationships[5].id = "rId5"

workbook.relationship_container.relationships.delete_if do | relationship |
  relationship.target == "worksheets/sheet#{del_sheet.sheet_id}.xml"
end

#add_sheet = workbook.add_worksheet
#add_sheet.sheet_name = 'Sheet1_Copy'
#add_sheet.sheet_data = template['Sheet1'].sheet_data

workbook.worksheets.delete(del_sheet)

# ブックの内容を列挙する
workbook.worksheets.each do |worksheet|
  p "シートのクラス名：#{worksheet.class.name}"
  p "シート名：#{worksheet.sheet_name}"
end

# 元から存在するシートを削除する
#delete_sheet_index = workbook.worksheets.index(org_sheet)
#workbook.worksheets.delete_at delete_sheet_index

# outputsディレクトリに保存する
file_name = "プレーンExcelサンプル削除_#{Time.now.strftime('%Y%m%d%H%M%s')}.xlsx"
workbook.write("/Users/kkanazaw/Dropbox/member_share/ssr/reports/#{file_name}")
#workbook.write("./#{file_name}")
