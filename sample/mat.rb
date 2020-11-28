require 'excelscan'
es = ExcelScan.new("sample/material.xlsx")

table = Hash.new(0)
es.each do |row|
  if row[0] then
    name = "#{row[1]} #{row[2]}"
    table[name] += row[3]
  end
end

es.quit

table.keys.sort.each do |name|
  printf "%-20s  %g\n", name, table[name]
end
