require 'rubygems'
require 'rspec'
require 'excel'
require 'kconv'

describe Excel do
		context "init and close" do
				it "normal end." do
						e = Excel.new("#{Dir.pwd}\\test.xlsx", "Sheet1")
						e.close()
				end
#				it "file not found end." do
#						e = Excel.new("#{Dir.pwd}\\aaa.xls")
#						e.close()
#				end
		end
		
		context "select SheetNames" do
				it "normal end." do
						e = Excel.new("#{Dir.pwd}\\test.xlsx", "Sheet1")
						e.getSheetNames
						e.close()
				end
		end
		
		context "select Values" do
				e = Excel.new("#{Dir.pwd}\\kanri.xls", Kconv.tosjis("宿題"))
				it "get headline." do
						e.getHeadline(1, 6)
				end
				it "get value." do
						e.getValues()
				end
				it "close." do
						e.close()
				end
		end
end
