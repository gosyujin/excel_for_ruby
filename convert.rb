require 'win32ole'
require 'kconv'
require 'csv'
require 'date'
require 'pp'

class Convert
  def initialize(read_cell_ini=nil)
      @Alpha = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 
                'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z']
      # Excelファイル定義
      @xls = nil
      @book = nil
      # 取得セルの定義ファイル
      @read_cell = ""
      if read_cell_ini.nil? then
        @read_cell = File::expand_path('.\read_cell')
      else
        @read_cell = File::expand_path(read_cell_ini)
      end
  end
  
  # Excelをオープンする
  def open(path)
    sheets = []
    begin
      @xls = WIN32OLE.new('Excel.Application')
      @book = @xls.Workbooks.Open(path)
      @book.Worksheets.each do |sheet|
        sheets << sheet.Name
      end
      return sheets
    rescue WIN32OLERuntimeError => ex
      #puts ex.message
      close()
    end
  end
  
  # Excelをクローズする
  def close()
    if @xls != nil then
      @xls.quit
      @xls = nil
    end
  end
  
  # CSVファイルを書き出す
  def output(path, value)
    unless File.exist?(path) then
      CSV.open(path, 'w')
    end
    f = File.open(path, 'a')
    CSV::Writer.generate(f) do |write|
      write << value
    end
    f.close()
  end
  
  # 特定のシートを取得する
  # sheet_name: 取得したいシート名
  #             nilの場合は一つ目(一番左)のシートを返す
  def get_sheet(sheet_name=nil)
    if sheet_name !=nil then
      return @book.Worksheets.Item(sheet_name)
    else
      @book.Worksheets.each do |sheet|
        return sheet
      end
    end
  end
  
  # 特定のセルの値を取得する
  # sheet: シート名
  # arg_cells: a1, A1 の場合 A1のセルを取得
  #            複数渡した場合はそれに対応する値を返す
  #            nilの場合は設定ファイルから抜き出す
  def get_value(sheet, arg_cells=nil)
    cells = []
    if arg_cells.nil? then
      f = File.open(@read_cell, 'r')
      f.each do |r|
        cells << r.downcase.chomp!
      end
    else
      arg_cells.each do |c|
        cells << c.downcase
      end
    end
    
    value = []
    cells.each do |cell|
      col, row = cell_to_address(cell)
      #puts "cells:#{cells}, x:#{col}, y:#{row}"
      value << sheet.Cells.Item(row, col).Value
    end
    return value 
  end
  
  # セルのアドレスとx, y座標を変換する
  # B3 ならば 2, 3となる
  def cell_to_address(cell)
    col = 0
    cell.scan(/[a-z]*/).each do |val|
      if val != "" then
        # x座標はBの場合は2、BBBの場合は
        # (2 * 26^2) + (2 * 26^1) + (2 * 26^0) = 1406
        inv = val.length - 1
        val.each_char do |c|
          col += (@Alpha.index(c) + 1) * (26 ** inv)
          inv -= 1
        end
      end
    end
    
    row = 0
    cell.scan(/[0-9]*/).each do |val|
      if val != "" then
        # y座標は抜き出した数字をそのまま使用する
        row = val.to_i
      end
    end
    return [col, row]
  end
end
