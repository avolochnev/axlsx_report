RSpec.shared_context 'with report' do |report_class|
  define_method :filename do
    '%s.xlsx' % report_class.name
  end

  define_method :report do
    @report ||= report_class.new
  end

  before :all do
    (1..10).each { |i| report << i }
    report.save(filename)
  end

  after :all do
    File.delete(filename)
    @report = nil
  end

  let :xlsx do
    Roo::Excelx.new(filename)
  end
end