require 'excelscan'

es = ExcelScan.new

col = es.insert
col[0] = "Table of 99"

9.times do |y|
  col = es.insert
  9.times do |x|
    col[x] = (x + 1) * (y + 1)
  end
end


