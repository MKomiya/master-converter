require 'roo'
require 'json'

# xlsx files to json
class Xlsx2Json
  def create(filepath)
    @s = Roo::Excelx.new(filepath)
    @filename = File.basename(filepath, '.xlsx')

    @headers = {}
    @hash = {}
  end

  def load_headers
    @s.sheets.each do |sheet|
      @s.default_sheet = sheet
      @headers[sheet] = {}
      (@s.first_column..@s.last_column).each do |col|
        @headers[sheet][col] = @s.cell(@s.first_row, col)
      end
    end
  end

  def load_data
    @s.sheets.each do |sheet|
      @s.default_sheet = sheet
      @hash[sheet] = []
      ((@s.first_row + 1)..@s.last_row).each do |row|
        @hash[sheet] << get_row_data(sheet, row)
      end
    end
  end

  def get_value(row, col)
    value = @s.cell(row, col)
    if @s.celltype(row, col) == :float && value.modulo(1) == 0.0
      value = value.to_i
    end
    value
  end

  def get_row_data(sheet, row)
    ret = {}
    @headers[sheet].keys.each do |col|
      ret[@headers[sheet][col]] = get_value(row, col)
    end
    ret
  end

  def run
    load_headers
    load_data

    puts @filename
    puts @hash.to_json
    File.write("master/#{@filename}.json", @hash.to_json)
  end
end
