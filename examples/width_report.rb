$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'axlsx_report'

class WidthReport < AxlsxReport::Base
  column 'Integer', width: 20 do |i|
    i
  end

  column 'Square', width: 20 do |i|
    i * i
  end
end

if __FILE__ == $0
  report = WidthReport.new
  (1..10).each { |i| report << i }
  report.save('width_report.xlsx')
end


