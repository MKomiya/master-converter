require 'roo'
require 'json'

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
        row_data = {}
        @headers[sheet].keys.each do |col|
          value = @s.cell(row, col)
          if @s.celltype(row, col) == :float && value.modulo(1) == 0.0
            value = value.to_i
          end
          row_data[@headers[sheet][col]] = value
        end
        @hash[sheet] << row_data
      end
    end
  end

  def run
    load_headers
    load_data

    puts @filename
    puts @hash.to_json
    File.write("master/#{@filename}.json", @hash.to_json)
  end
end
