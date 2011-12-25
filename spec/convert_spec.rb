require 'rubygems'
require 'rspec'
require 'convert'
require 'kconv'
require 'csv'
require 'pp'

describe Convert do
# プラットフォームがWindowsの場合標準出力をKconv.tosjisでラップ
if RUBY_PLATFORM.downcase =~ /mswin(?!ce)|mingw|cygwin|bccwin/ then
  def $stdout.write(str)
    super Kconv.tosjis(str)
  end
end
  before do
    @c = Convert.new()
    # input file
    @xls_path = File::expand_path('.\testdata\input\input.xlsx')
    # output file
    @csv_path = File::expand_path('.\testdata\output\output.csv')
    @csv_path2 = File::expand_path('.\testdata\output\output2.csv')
    
    @first_sheet_name = Kconv.tosjis("レポート")
    @a1_value = Kconv.tosjis("TEST")
    
    @cell_value = Hash.new
    @cell_value["B2"] = Kconv.tosjis("ここは見出し")
    @cell_value["B3"] = Kconv.tosjis("件名")
    @cell_value["C3"] = Kconv.tosjis("件名ほげほげ")
    @cell_value["B4"] = Kconv.tosjis("No")
    @cell_value["C4"] = Kconv.tosjis("55")
    @cell_value["C6"] = Kconv.tosjis("2011/12/08 00:00:00")
    # エスケープシーケンス(改行)注意
    @cell_value["C8"] = Kconv.tosjis("備考\n　　備考")
    @cell_value["BBB300"] = "CLEAR!!"
  end
  describe 'Excelファイルをオープンする時' do
    it '正常にオープンできる' do
      book = @c.open(@xls_path)
      book[0].should be == @first_sheet_name
    end
    it '特定のシートオブジェクトを取得できる' do
      book = @c.open(@xls_path)
      sheet = @c.get_sheet(book[0])
      sheet.Name.should be == @first_sheet_name
    end
    it '1枚目のシートを取得でき、中のA1セルの値を取得できる' do
      book = @c.open(@xls_path)
      sheet = @c.get_sheet()
      sheet.Name.should be == @first_sheet_name
      sheet.Cells.Item(1, 1).Value.should be == @a1_value
    end
    it 'セルを指定したら該当セルの配列を取得できる' do
      sheet = @c.get_sheet(@c.open(@xls_path)[0])
      @cell_value.each do |k, v|
        @c.get_value(sheet, k).should be == [v]
      end
    end
    it '複数セルを指定したら該当セルの配列を取得できる' do
      sheet = @c.get_sheet(@c.open(@xls_path)[0])
      @c.get_value(sheet, ["B4", "C4"]).should be == 
        [@cell_value["B4"], @cell_value["C4"]]
      @c.get_value(sheet, ["B3", "C3"]).should be == 
        [@cell_value["B3"], @cell_value["C3"]]
    end

    # 一つにまとめられそう
    describe 'CSVファイルに書き出す時' do
      before do 
        begin
          if File.exist?(@csv_path) then
            File.delete(@csv_path)
          end
        rescue => ex
          # Permission denied(使用中)で消せない場合
          # 例外を投げるのでPendingする
          pending(ex)
        end
      end
      it '正常に1レコード書き出せる(新規作成)' do
        sheet = @c.get_sheet(@c.open(@xls_path)[0])
        cell_values = @c.get_value(sheet, ["B2", "B3", "C3", "B4", "C4"])
        @c.output(@csv_path, cell_values)
        @rows = []
        CSV.open(@csv_path, 'r') do |row|
          @rows << row
        end
        @rows[0].should be == cell_values
      end
#      it '正常に書き出せる(追記あり)' do
#      end
    end
    describe 'CSVファイルに書き出す時' do
      before do 
        begin
          if File.exist?(@csv_path2) then
            File.delete(@csv_path2)
          end
        rescue => ex
          # Permission denied(使用中)で消せない場合
          # 例外を投げるのでPendingする
          pending(ex)
        end
      end
      it '取得するシートを外ファイルから読み出せる' do
        @c = Convert.new(File::expand_path('.\testdata\init\read_cell'))
        sheet = @c.get_sheet(@c.open(@xls_path)[0])
        cell_values = @c.get_value(sheet)
        @c.output(@csv_path2, cell_values)
        @rows = []
        CSV.open(@csv_path2, 'r') do |row|
          @rows << row
        end
        @rows[0].should be == cell_values
      end
    end
  end
  after do
    @c.close()
  end
end
