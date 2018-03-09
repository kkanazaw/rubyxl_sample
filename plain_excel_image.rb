# coding: utf-8
require 'rubyXL'

# ワークブックの読込
#template = RubyXL::Parser.parse('excels/プレーンExcelサンプル.xlsx')
#workbook = RubyXL::Parser.parse('excels/プレーンExcelサンプル.xlsx')

template = RubyXL::Parser.parse('excels/hand_copy2.xlsx')
workbook = RubyXL::Parser.parse('excels/hand_copy2.xlsx')

# コピー元シート
add_sheet =  template['その2']

#add_sheet = Marshal.load(Marshal.dump(org_sheet))

#add_sheet = org_sheet
add_sheet.sheet_name = 'その2_Copy'
add_sheet.workbook = workbook

# worksheetの番号を取得する
add_sheet.sheet_id = workbook.worksheets.map(&:sheet_id).max + 1

#add_sheet.printer_settings[0].xlsx_path = ::Pathname.new("/xl/printerSettings/printerSettings#{add_sheet.sheet_id}.bin")
add_sheet.generic_storage[0].xlsx_path = ::Pathname.new("/xl/drawings/drawing#{add_sheet.sheet_id}.xml")

add_sheet.relationship_container.relationships[0].target = "../drawings/drawing#{add_sheet.sheet_id}.xml"
                                                                          
#  .xslx_path = ::Pathname.new("/xl/printerSettings/printerSettings#{add_sheet.sheet_id}.bin")

# テストシートをコピーしてブックに追加
workbook.worksheets << add_sheet 

workbook.relationship_container.relationships << RubyXL::Relationship.new(
  :id => "rId#{workbook.relationship_container.relationships.size + 1}",
  :type => add_sheet.class::REL_TYPE,
  :target => "worksheets/sheet#{add_sheet.sheet_id}.xml")

#add_sheet = workbook.add_worksheet
#add_sheet.sheet_name = 'Sheet1_Copy'
#add_sheet.sheet_data = template['Sheet1'].sheet_data


# ブックの内容を列挙する
workbook.worksheets.each do |worksheet|
  p "シートのクラス名：#{worksheet.class.name}"
  p "シート名：#{worksheet.sheet_name}"
end

# 元から存在するシートを削除する
#delete_sheet_index = workbook.worksheets.index(org_sheet)
#workbook.worksheets.delete_at delete_sheet_index

# outputsディレクトリに保存する
file_name = "プレーンExcelサンプル_画像_#{Time.now.strftime('%Y%m%d%H%M%s')}.xlsx"
workbook.write("/Users/kkanazaw/Dropbox/member_share/ssr/reports/#{file_name}")
#workbook.write("./#{file_name}")
