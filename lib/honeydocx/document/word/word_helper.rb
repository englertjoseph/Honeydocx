module Honeydocx
  module WordHelper
    def read_fixture(filename)
      File.open(fixture_file(filename)).read
    end

    def fixture_file(filename)
      File.expand_path("../word_fixtures/#{filename}" , __FILE__)
    end

    def template_header_rels
      File.open(WordXML.honey_header_rels_path).read
    end

    #insert files?
  end
end
