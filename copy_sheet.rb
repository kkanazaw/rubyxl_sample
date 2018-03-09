# coding: utf-8
require 'zip'
require 'nokogiri'
require 'rubyXL'

# def parse_file(zip_file, file_path)
#   entry = zip_file.find_entry(file_path)
#   entry && (entry.get_input_stream { |f| f.read })
# end
    

  stream = Zip::OutputStream.write_buffer { |zipstream |
    xml_string = write_xml
    parsed.each do | key, value|
    end
    zipstream.put_next_entry(self.xlsx_path)
    zipstream.write(xml_string)
  }
  stream.rewind
  RubyXL::Workbook::arse_buffer(stream)
  
    
begin
  src_file_path = "excels/プレーンExcelサンプル.xlsx"
  parsed = {}
  ::Zip::File.open(src_file_path) do |zip_file|
    zip_file.each do | entry |
      parsed[entry.name] = Nokogiri::XML(entry.get_input_stream.read)
    end
  end
  
  sheets = parsed["xl/worksheets/_rels/sheet4.xml.rels"] = parsed["xl/worksheets/_rels/sheet4.xml.rels"]

  p parsed["xl/workbook.xml"].to_xml
rescue ::Zip::Error => e
  raise e, "XLSX file format error: #{e}", e.backtrace
end

