$LOAD_PATH.unshift File.expand_path('../../lib', __FILE__)
require 'honeydocx'

def get_header(document)
  header = open_xml('word/header1.xml', document)
end

# Remove spaces and rsids from the header for comparisons.
def clean_xml(header)
  header.gsub(/\s+|w:rsid.*?"[^\"]*"|standalone.*"/, "")
end

def expected_header
  header = File.open(File.expand_path(
    '../fixtures/header1.xml', __FILE__)).read
end

private

def open_xml(filename, path)
  Zip::File.open(path).read(filename)
end
