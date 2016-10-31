require 'spec_helper'
require './examples/lambda_report'
require 'roo'

describe LambdaReport do
  include_context 'with report', LambdaReport

  it 'have 11 rows' do
    expect(xlsx.sheet(0).last_row).to eq 11
  end

 it 'use lambda to calculate value' do
    expect(xlsx.sheet(0).cell('B', 6)).to eq true
    expect(xlsx.sheet(0).cell('C', 6)).to eq false
  end
end
