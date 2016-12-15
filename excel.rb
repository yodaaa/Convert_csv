# coding: utf-8
# 2015/08/08
require 'csv'
require 'spreadsheet'

Spreadsheet.client_encoding = 'UTF-8'
book = Spreadsheet.open 'te.xls' #ファイルの指定
sheet1 = book.worksheet 0        #シートの指定
count = 0

file = open("file.csv", "w")
#Excelシートの読み込み
sheet1.each 4805 do |row|
  a = []
  row[0] = "#{row[0]}".gsub(/-/, '/').gsub(/T/, ' ').gsub(/\+00:00/, '')
  a <<  "#{row[0]},#{row[1]},#{row[2]},#{row[3]},#{row[4].instance_of?(Spreadsheet::Excel::Error) ? "0.0" : row[4]},#{row[5].instance_of?(Spreadsheet::Excel::Error) ? "0.0" : row[5]}"
  puts f = a.to_s
  file.write(f.gsub(/\[\"/, '').gsub(/\"\]/, ''))
  file.write("\n")
=begin
      count += 1
      if count > 1000 then #データが大量なため、一部抜粋する場合
        break
      end
=end
end
