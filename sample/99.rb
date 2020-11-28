# -*- coding: cp932 -*-
require 'excelscan'

es = ExcelScan.new(nil)

col = es.insert
col[0] = "９９の表"

9.times do |y|
  col = es.insert
  9.times do |x|
    col[x] = (x + 1) * (y + 1)
  end
end


