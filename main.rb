require 'rubyXL'

# ワークブックの読込
workbook = RubyXL::Parser.parse('excels/beaglesoft_template2.xlsx')

# note:こちらは動作しない
# add_sheet = Marshal.load(Marshal.dump(workbook['テストシート名']))
# add_sheet.sheet_name = 'コピーシート名'
# add_sheet.workbook = workbook2

# テストシートをコピーしてブックに追加
# workbook2.worksheets << add_sheet

worksheet = workbook.add_worksheet
worksheet.sheet_data = workbook['テストシート名'].sheet_data

# ブックの内容を列挙する
workbook.worksheets.each do |worksheet|
  p "シートのクラス名：#{worksheet.class.name}"
  p "シート名：#{worksheet.sheet_name}"
end

# 元から存在するシートを削除する
# workbook.worksheets.delete_at(0)

# outputsディレクトリに保存する
file_name = "beaglesoft_template_#{Time.now.strftime('%Y%m%d%H%M%s')}.xlsx"
workbook.write("outputs/#{file_name}")
