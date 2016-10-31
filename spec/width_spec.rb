require 'spec_helper'
require './examples/width_report'
require 'roo'

describe WidthReport do
  include_context 'with report', WidthReport

  it 'have 11 rows' do
    expect(xlsx.sheet(0).last_row).to eq 11
  end

  it 'have column width setup' do
    expect(report.sheet.send(:find_or_create_column_info, 0).width).to eq 20
  end
end
