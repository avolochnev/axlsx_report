$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'axlsx_report'

class BaseReport < AxlsxReport::Base
  column 'Integer' do |i|
    i
  end

  column 'Square' do |i|
    i * i
  end
end

if __FILE__ == $0
  report = BaseReport.new
  (1..10).each { |i| report << i }
  report.save('base_report.xlsx')
end


