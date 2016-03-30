Gem::Specification.new do |s|
  s.name        = 'quick_spreadsheet'
  s.version     = '0.1.0'
  s.date        = '2016-03-30'
  s.summary     = "Wraps other utilities to provide sensible defaults for spreadsheet generation."
  s.description = s.summary
  s.authors     = ["Brian Gracie","Daniel Museakamp"]
  s.email       = 'bgracie@gmail.com'
  s.files       = ["lib/quick_spreadsheet.rb"]
  s.license     = 'MIT'

  s.add_runtime_dependency "spreadsheet",
    "~> 1.0.9"
end
