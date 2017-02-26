
require 'spreadsheet'

# ============================================
# ===========   Read Example   ===============
# ============================================

# Note: spreadsheet only supports .xls files (not .xlsx)
workbook = Spreadsheet.open './sample_excel_files/sample_xls_excel_file.xls'

worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.rows.each do |row|
    row_cells = row.to_a.map{ |v| v.methods.include?(:value) ? v.value : v }
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join " "
  end
  puts "Read #{num_rows} rows"
end

puts 'Done'
