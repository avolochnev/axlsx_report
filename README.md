# AxlsxReport

Declarative excel reports based on [axlsx](https://github.com/caxlsx/caxlsx).

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'axlsx_report'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install axlsx_report

## Usage

```ruby
require 'axlsx_report'

class BaseReport < AxlsxReport::Base
  column 'Number' do |i|
    i
  end

  column 'Square' do |i|
    i * i
  end
end

report = BaseReport.new

(1..10).each do |i|
  report << i
end

report.save('base_report.xlsx')
```

Find more examples in examples folder.

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/avolochnev/axlsx_report.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

