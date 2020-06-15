require 'robust_excel_ole'

# ============================================
# ===========   Read Example   ===============
# ============================================

workbook = RobustExcelOle::Workbook.open './sample_excel_files/xlsx_500_rows.xlsx'

puts "Found #{workbook.worksheets_count} worksheets"

workbook.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.values.each do |row_vals|
    row_cells = row_vals
    num_rows += 1
  end

  puts "Read #{num_rows} rows"

end

puts 'Done'
