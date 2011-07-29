require 'win32ole'
require 'kconv'
require 'date'

# = Excelファイルを操作するクラス
class Excel
		# 初期化
		def initialize(file, sheetName)
				begin
						@xls = WIN32OLE.new('Excel.Application')
						# Excelファイルを開く
						@book = @xls.Workbooks.Open(file)
						getSheetName(sheetName)
				rescue WIN32OLERuntimeError
						putsError("File not found. - #{file}")
						close()
						exit
				end
		end
		
		# 決め打ちでほしいデータを持ってきておる
		def getValues(x=1, y=7)
				# A列の値(ID,Noを想定)がある限りループ
				while @sheet.Cells.Item(y, 1).Value != nil
						value = ""
						
						# K列の値(状況のステータスを想定)を比較
						# 完了していないものを抽出
						if @sheet.Cells.Item(y, 11).Value != Kconv.tosjis("完了") then
								# J列の値(日限を想定)を比較
								# 期限が直近(一週間後)まで、または
								# 期限が定められていない(nilものを抽出
								if @sheet.Cells.Item(y, 10).Value == nil or
										Date.strptime(
												@sheet.Cells.Item(y, 10).Value, 
												"%Y/%m/%d %H:%M:%S"
										) <= Date.today + 7 then
												# http://cyakarin.kuronowish.com/index.cgi?tof
												# No.[項番]:
												number = @sheet.Cells.Item(y, 1).Value || "0"
												value += "No.#{number.to_i}:"
												# ([担当者])
												charge = @sheet.Cells.Item(y, 9).Value || ""
												value += "(#{charge.gsub!("\n", "")})"
												# [内容]\n
												content = @sheet.Cells.Item(y, 5).Value || ""
												value += "#{content}\n"
												# [現在の状況]
												progress = @sheet.Cells.Item(y, 7).Value || ""
												value += "-> #{progress}\n"
												
												value += "---\n"
												
												puts value
								end
						end
						y += 1
				end
		end
		
		# シート名一覧を取得する
		def getSheetNames
				sheets = []
				@book.Worksheets.each do |sheet|
						sheets << sheet.Name
				end
				puts sheets.join(",")
		end
		
		# シート名を取得する
		def getSheetName(sheetName)
				begin
						@sheet = @book.Worksheets.Item(sheetName)
				rescue WIN32OLERuntimeError
						putsError("Sheet not found. - #{sheetName} ")
						close()
						exit
				end
		end
		
		# 見出しとなる行を取得する
		# スタート地点(x, y)から x + 1 方向へ進み
		# Nilが返ってくるまでの値を配列に代入
		def getHeadline(x, y)
				@headline = []
				while @sheet.Cells.Item(y, x).Value != nil
						@headline << @sheet.Cells.Item(y, x).Value
						x += 1
				end
		end
		
		# シートの全値を取得する
		def selectSheetValue(sheetName)
						@sheet.UsedRange.Rows.each do |row|
								record = []
								row.Columns.each do |cell|
										if cell.Value != nil then
												record << cell.Value
										end
								end
								puts record.join(",")
						end
		end
		
		# Excelファイルをクローズする
		def close
#				puts "close"
				@book.Close
				@xls.quit
		end
		
		# プラットフォームがWindowsの場合標準出力をKconv.tosjisでラップ
		if RUBY_PLATFORM.downcase =~ /mswin(?!ce)|mingw|cygwin|bccwin/ then
				def $stdout.write(str)
						super Kconv.tosjis(str)
				end
		end

		# エラーの起きたメソッドを出力する
		# http://www.rubyist.net/~nobu/t/20051013.html#p02
		def putsError(ex="")
				puts "[Error]" + caller.first[/:in \`(.*?)\'\z/, 1] + " - " + ex
		end
end
