require 'spreadsheet'
require 'time'

class QuickSpreadsheet
  def self.args
    puts "title: string"
    puts "header_row: Array<string>"
    puts "body_rows: Array<Array<string>>"
    puts "folder_path: string (optional)"
  end

  def self.call(args)
    title       = args[:title]
    header_row  = args[:header_row]
    body_rows   = args[:body_rows]
    folder_path = args[:folder_path] || "#{Dir.pwd}/tmp"
    format      = :xls

    FileUtils::mkdir_p(folder_path) unless File.exists?(folder_path)

    time = Time.now
    file_timestamp = time.strftime("%Y-%m-%d-%H.%M")
    filename = "#{title.gsub(" ", "")}_g#{file_timestamp}.#{format.to_s}"
    file_path = [ folder_path, filename ].join("/")
    content_timestamp = time.strftime("%Y-%m-%d-%H:%M")
    rows = [
      [ title ],
      [ "Generated: #{content_timestamp}" ],
      [ "" ],
      header_row
    ] + body_rows

    WRITE_FORMAT[format].call(rows, title, file_path)

    puts "Wrote your spreadsheet to '#{file_path}'."

    true
  end

  WRITE_FORMAT = {
    xls: Proc.new do |rows, title, file_path|
      book = Spreadsheet::Workbook.new
      sheet = book.create_worksheet name: title

      rows.each.with_index do |row, index|
        sheet.row(index).concat(row)
      end

      book.write file_path
    end
  }
end
