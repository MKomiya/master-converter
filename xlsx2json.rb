require 'roo'
require 'json'

begin
    s = Roo::Excelx.new($*[0])
rescue
    STDERR.puts "Usage: ruby xlsx2json.rb excelx_file"
    exit false
end

s.default_sheet = s.sheets.first

headers = {}
(s.first_column..s.last_column).each do |col|
    headers[col] = s.cell(s.first_row, col)
end

hash = {}
hash[s.default_sheet] = []
((s.first_row + 1)..s.last_row).each do |row|
    row_data = {}
    headers.keys.each do |col|
        value = s.cell(row, col)
        value = value.to_i if s.celltype(row,col) == :float && value.modulo(1) == 0.0
        row_data[headers[col]] = value
    end
    hash[s.default_sheet] << row_data
end

puts hash.to_json
File.write('master/item.json', hash.to_json)
