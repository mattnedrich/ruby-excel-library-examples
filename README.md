# ruby-excel-library-examples
This project contains sample code for reading Excel files with different Ruby libraries.

## `.xlsx` File Examples
Below are code samples for reading current OOXML Excel files using [**rubyXL**](https://github.com/weshatheleopard/rubyXL), [**roo**](https://github.com/roo-rb/roo), [**creek**](https://github.com/pythonicrubyist/creek), and [**simple_xlsx_reader**](https://github.com/woahdae/simple_xlsx_reader).

### [rubyXL](https://github.com/weshatheleopard/rubyXL)
```ruby
require 'rubyXL'

workbook = RubyXL::Parser.parse './sample_excel_files/xlsx_500_rows.xlsx'
worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.sheet_name}"
  num_rows = 0

  worksheet.each do |row|
    row_cells = row.cells.map{ |cell| cell.value }
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join " "
  end
  puts "Read #{num_rows} rows"
end

puts 'Done'
```
### [roo](https://github.com/roo-rb/roo)
```ruby
require 'roo'

workbook = Roo::Spreadsheet.open './sample_excel_files/xlsx_500_rows.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet}"
  num_rows = 0

  workbook.sheet(worksheet).each_row_streaming do |row|
    row_cells = row.map { |cell| cell.value }
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join ' '
  end
  puts "Read #{num_rows} rows"
end

puts 'Done'
```
### [creek](https://github.com/pythonicrubyist/creek)
```ruby
require 'creek'

workbook = Creek::Book.new './sample_excel_files/xlsx_500_rows.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.name}"
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
```
### [simple_xlsx_reader](https://github.com/woahdae/simple_xlsx_reader)
```ruby
require 'simple_xlsx_reader'

workbook = SimpleXlsxReader.open './sample_excel_files/xlsx_500000_rows.xlsx'
worksheets = workbook.sheets
puts "Found #{worksheets.count} worksheets"

worksheets.each do |worksheet|
  puts "Reading: #{worksheet.name}"
  num_rows = 0

  worksheet.rows.each do |row|
    row_cells = row
    num_rows += 1

    # uncomment to print out row values
    # puts row_cells.join ' '
  end
  puts "Read #{num_rows} rows"
end

puts 'Done'
```

## Legacy `.xls` Files
Below are code samples for reading legacy Excel files using [**spreadsheet**](https://github.com/zdavatz/spreadsheet)

### [spreadsheet](https://github.com/zdavatz/spreadsheet)
```ruby
require 'spreadsheet'

# Note: spreadsheet only supports .xls files (not .xlsx)
workbook = Spreadsheet.open './sample_excel_files/xls_500_rows.xls'
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
```
