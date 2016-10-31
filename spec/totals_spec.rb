require 'spec_helper'
require './examples/totals_report'
require 'roo'

describe TotalsReport do
  include_context 'with report', TotalsReport

  it 'have 12 rows' do
    expect(xlsx.sheet(0).last_row).to eq 12
  end

  it 'have valid total actions' do
    expect(report.send(:total_actions)).to eq ['Total', :sum]
  end

  it 'display string total' do
    expect(xlsx.sheet(0).cell('A', 12)).to eq 'Total'
  end

  it 'calculate sum total' do
    expect(xlsx.sheet(0).cell('B', 12)).to eq 385
  end
end
