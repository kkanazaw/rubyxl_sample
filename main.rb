require 'rubyXL'

# ワークブックの読込
workbook = RubyXL::Parser.parse('excels/beaglesoft_template.xlsx')
add_sheet = workbook['テストシート名'].dup
add_sheet.sheet_name = 'コピーシート名'

# テストシートをコピーしてブックに追加
workbook.worksheets << add_sheet

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
