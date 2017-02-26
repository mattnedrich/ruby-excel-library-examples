
require 'creek'

# ============================================
# ===========   Read Example   ===============
# ============================================

workbook = Creek::Book.new "sample_excel_file.xlsx"

worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet}"
  num_rows = 0

  worksheet.rows.each do |row|
    row_cells = row.values
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join " "
  end
  puts "Read #{num_rows} rows"
end

puts 'Done'
