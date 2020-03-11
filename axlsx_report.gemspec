# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'axlsx_report/version'

Gem::Specification.new do |spec|
  spec.name          = "axlsx_report"
  spec.version       = AxlsxReport::VERSION
  spec.authors       = ["Alexey Volochnev"]
  spec.email         = ["alexey.volochnev@gmail.com"]

  spec.summary       = %q{Declarative excel reports based on axlsx.}
  spec.homepage      = "https://github.com/avolochnev/axlsx_report"
  spec.license       = "MIT"

  spec.files         = `git ls-files -z`.split("\x0").reject { |f| f.match(%r{^(test|spec|features)/}) }
  spec.require_paths = ["lib"]

  spec.add_runtime_dependency 'caxlsx', '~> 3.0', '>= 3.0.1'
  spec.add_development_dependency "bundler", "~> 1.11"
  spec.add_development_dependency "rake", ">= 12.3.3"
  spec.add_development_dependency "rspec", "~> 3.0"
  spec.add_development_dependency 'roo', '~> 2.8', '>= 2.8.3'
end
