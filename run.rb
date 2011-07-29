require 'excel'

if ARGV.length != 2 then
		puts "Usage: #{$0} EXCEL_FILE SHEET_NAME"
		exit
end

e = Excel.new(ARGV[0], ARGV[1])
e.getValues
