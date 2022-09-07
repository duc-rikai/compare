require 'rubyXL'
require 'rubyXL/convenience_methods'

p "START READ FILE"
workbook = RubyXL::Parser.parse './Contact Rikai.xlsx'
old_data = workbook['None IT Companies']
new_data = workbook['Non IT 0829']
p "END READ FILE"

array_old_data = []
array_new_data = []
data_new = []
puts "START GET DATA SHEET 1"
(1...old_data.to_a.length).each do |i|
  array_old_data << old_data[i][3].value unless old_data[i][3].nil?
end
puts "GET DATA SHEET 1 SUCCESSFUL"

puts "START GET DATA SHEET 2"
(1...new_data.to_a.length).each do |i|
  array_new_data << new_data[i][3].value unless new_data[i][3].nil?
end
puts "GET DATA SHEET 2 SUCCESSFUL"

array_dup = array_old_data & array_new_data
sheet_array = (array_old_data + array_new_data) - array_dup
new_array = array_new_data & sheet_array

workbook_new = RubyXL::Workbook.new [new_data]
worksheet = workbook_new[0]
worksheet.add_cell(0, 0, new_data[0][0].value)
worksheet.add_cell(0, 1, new_data[0][1].value)
worksheet.add_cell(0, 2, new_data[0][2].value)
worksheet.add_cell(0, 3, new_data[0][3].value)
worksheet.add_cell(0, 4, new_data[0][4].value)
worksheet.add_cell(0, 5, new_data[0][5].value)
p "START"

c = 0
(1...new_data.to_a.length).each do |i|
  value = new_array & new_data[i][3].value.split(' ')
  if value.length > 0
    c += 1
    new_data[i][0].nil? ? next : worksheet.add_cell(c, 0, new_data[i][0].value)
    new_data[i][1].nil? ? next : worksheet.add_cell(c, 1, new_data[i][1].value)
    new_data[i][2].nil? ? next : worksheet.add_cell(c, 2, new_data[i][2].value)
    new_data[i][3].nil? ? next : worksheet.add_cell(c, 3, new_data[i][3].value)
    new_data[i][4].nil? ? next : worksheet.add_cell(c, 4, new_data[i][4].value)
    new_data[i][5].nil? ? next : worksheet.add_cell(c, 5, new_data[i][5].value)
  end
end
workbook_new.write './Compare.xlsx'
p "DONE"
