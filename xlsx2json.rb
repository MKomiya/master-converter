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
s.sheets.each do |sheet|
    s.default_sheet = sheet
    headers[sheet] = {}
    (s.first_column..s.last_column).each do |col|
        headers[sheet][col] = s.cell(s.first_row, col)
    end
end

hash = {}
s.sheets.each do |sheet|
    s.default_sheet = sheet
    hash[sheet] = [] 
    ((s.first_row + 1)..s.last_row).each do |row|
        row_data = {}
        headers[sheet].keys.each do |col|
            value = s.cell(row, col)
            value = value.to_i if s.celltype(row,col) == :float && value.modulo(1) == 0.0
            row_data[headers[sheet][col]] = value
        end
        hash[sheet] << row_data
    end
end

puts hash.to_json
File.write('master/master.json', hash.to_json)
