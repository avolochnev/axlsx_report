require 'spec_helper'
require './examples/base_report'
require 'roo'

describe BaseReport do
  include_context 'with report', BaseReport

  it 'have 11 rows' do
    expect(xlsx.sheet(0).last_row).to eq 11
  end

  it 'use formula to calculate value' do
    expect(xlsx.sheet(0).cell('B', 10)).to eq 81
  end

  it 'store headers' do
    expect(xlsx.sheet(0).row(1)).to eq ['Integer', 'Square']
  end
end
