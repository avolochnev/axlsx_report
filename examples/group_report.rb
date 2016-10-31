$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../lib"

require 'axlsx_report'

class GroupReport < AxlsxReport::Base
  column 'Integer' do |i|
    i
  end

  # Define group of columns with merged group header
  group 'Calculations' do
    column 'Square' do |i|
      i * i
    end

    column 'Sqrt' do |i|
      Math.sqrt(i)
    end
  end
end

if __FILE__ == $0
  report = GroupReport.new
  (1..10).each { |i| report << i }
  report.save('group_report.xlsx')
end


