require 'optparse'
require 'convert'
require 'kconv'
require 'date'
require 'pp'

c = Convert.new()
mode = Hash.new
OptionParser.new do |o|
  o.on_head('-i VALUE') do |v|
    mode['i'] = v
  end
  o.on_head('-o VALUE') do |v|
    mode['o'] = v
  end
  o.parse!(ARGV)
end

if mode['i'] && mode['o'] then
  puts "modei #{mode['i']}"
  puts "modeo #{mode['o']}"
  sheet = c.get_sheet(c.open(mode['i'])[0])
  cell_values = c.get_value(sheet)
  c.output(mode['o'], cell_values)
else
  puts "Usage nyoronyoro."
end
c.close()
