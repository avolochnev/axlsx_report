require 'spec_helper'
require './examples/group_report'
require 'roo'

describe GroupReport do
  include_context 'with report', GroupReport

  it 'have 12 rows' do
    expect(xlsx.sheet(0).last_row).to eq 12
  end

  it 'store group headers' do
    # Roo do not allow to verify if the cell is merged. It is just return the same value for all cells in merged group.
    expect(xlsx.sheet(0).row(1)).to eq ['Integer', 'Calculations', 'Sqrt']
  end

  it 'have header merged' do
    expect(report.sheet.send(:merged_cells).to_a).to eq ["B1:C1", "A1:A2"]
  end

  it 'store column headers' do
    expect(xlsx.sheet(0).row(2)).to eq ['Integer', 'Square', 'Sqrt']
  end
end
