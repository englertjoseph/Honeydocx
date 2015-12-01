

module Honeydocx
  class Document
    def create(type, opts={})
      if type == :pdf
        return PDF.new(opts)
      elsif type == :docx
        return WordXML.new(opts)
      elsif type == :xlsx
        return ExcelXML.new(opts)
      end
    end

    def add_honey

    end

    def save!
      fail
    end

    def self.accepted_formats
      ['.pdf', '.docx', '.xlsx']
    end

    def self.parse_filetype(filename)
      ext = File.extname(filename)
      accepted_formats.include?(ext) ? ext : raise(UnknownFormatError, filename)
    end

    class UnknownFormatError < StandardError
      def initialize(format)
        super("Error the format #{format} is not understood.")
      end
    end
  end
end

require_relative 'word_xml'
require_relative 'excel_xml'
require_relative 'pdf'
