module Honeydocx
  class WordXML < Document

    attr_reader :path, :zip, :doc_xml, :doc

    def initialize(path, opts={})
      @filename = path
      open_document
    end

    def add_honey

    end

    def save!

    end

    private

    def open_document
      @zip = Zip::File.open(path)
      @doc_xml = @zip.read('word/document.xml')
      #@doc = Nokogiri::XML(doc_xml)
      binding.pry
    end
  end
end
