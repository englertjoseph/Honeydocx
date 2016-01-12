module Honeydocx
  module WordHelper
    def read_fixture(filename)
      File.open(fixture_file(filename)).read
    end

    def fixture_file(filename)
      File.expand_path("../../word_fixtures/#{filename}" , __FILE__)
    end

    def open_xml(file)
      Nokogiri::XML(file)
    end

    def template_header_rels
      File.open(WordXML.honey_header_rels_path).read
    end

    def has_header_rels?
      doc.zip.entries.any? { |entry| entry.name == "word/_rels/header1.xml.rels" }
    end

    def add_file_to_zip(filename, data)
      doc.add_file_to_zip(filename, data)
    end

    #insert files?
  end
end
