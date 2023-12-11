require 'spreadsheet'
require 'time'

class QuickSpreadsheet
  def self.args
    puts "file_title: string (optional)"
    puts "folder_path: string (optional)"
    puts "Variant 1:"
    puts "  title: string"
    puts "  header_row: Array<string>"
    puts "  body_rows: Array<Array<string>>"
    puts "Variant 2:"
    puts "  sheets: Array<Array<"
    puts "    title: string"
    puts "    header_row: Array<string>"
    puts "    body_rows: Array<Array<string>>"
  end

  def self.call(args)
    time = Time.now

    file_timestamp = time.strftime("%Y-%m-%d-%H.%M")

    folder_path = args[:folder_path] || "#{Dir.pwd}/tmp"

    format = :xls

    sheets =
      if args[:sheets]
        args[:sheets]
      else
        [{
          title:      args[:title],
          header_row: args[:header_row],
          body_rows:  args[:body_rows]
        }]
      end
    
    filename =
      if args[:file_title].nil?
        "#{sheets[0][:title].gsub(" ", "")}_g#{file_timestamp}.#{format.to_s}"
      else
        "#{args[:file_title].gsub(" ", "")}_g#{file_timestamp}.#{format.to_s}"
      end

    file_path = [ folder_path, filename ].join("/")

    FileUtils::mkdir_p(folder_path) unless File.exist?(folder_path)

    WRITE_FORMAT[format].call(sheets, file_path)

    puts "Wrote your spreadsheet to '#{file_path}'."

    file_path
  end

  WRITE_FORMAT = {
    xls: Proc.new do |sheets, file_path|
      book = Spreadsheet::Workbook.new
      sheets.each do |sheet|
        worksheet = book.create_worksheet name: sheet[:title]
        rows = [
          sheet[:header_row]
        ] + sheet[:body_rows]
        rows.each.with_index do |row, index|
          worksheet.row(index).concat(row)
        end
      end
      book.write file_path
    end
  }
end

# There are two ways to call, and argument order is not important.
# In each syntax, file_title and folder_path are optional. Defaults to a
# file_title from the first (or only) title in a local tmp folder.

# Call like this:

# QuickSpreadsheet.call(
#   file_title: "Some Name",
#   folder_path: "some/folder/path",
#   title: "My Spreadsheet",
#   header_row: ["Some Header 1","Some Header 2"],
#   body_rows: [['1','a'],['2','b']]
# )

# Or like this:

# QuickSpreadsheet.call(
#   file_title: "Some Name",
#   folder_path: "some/folder/path",
#   sheets: [
#     {
#       title: "My First Sheet",
#       header_row: ["Some Header 1","Header 2","Header 3"],
#       body_rows: [['1','a','i'],['2','b','ii'],['3','c','iii']]
#     },{
#       title: "My Second Sheet",
#       header_row: ["Some Other Header 1","Other Header 2","Other Header 3"],
#       body_rows: [['10','ax','xi'],['20','bx','xii'],['30','cx','xiii']]
#     }
#   ]
# )
