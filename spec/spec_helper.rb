$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'axlsx_report'
Dir["./spec/support/**/*.rb"].each {|f| require f}
