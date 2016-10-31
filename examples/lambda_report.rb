$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'axlsx_report'

class LambdaReport < AxlsxReport::Base
  # proc without params called in context of given object
  column 'Integer', -> { self }
  column 'Odd?', -> { odd? }, width: 7
  column 'Even?', -> { even? }, width: 7
end

if __FILE__ == $0
  report = LambdaReport.new
  (1..10).each { |i| report << i }
  report.save('lambda_report.xlsx')
end


