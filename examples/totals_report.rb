$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'axlsx_report'

class TotalsReport < AxlsxReport::Base
  column 'Integer', ->(i) { i },     total: 'Total'
  column 'Square',  ->(i) { i * i }, total: :sum
end

if __FILE__ == $0
  report = TotalsReport.new
  (1..10).each { |i| report << i }
  report.save('totals_report.xlsx')
end


