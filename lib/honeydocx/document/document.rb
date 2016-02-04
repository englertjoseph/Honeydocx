module Honeydocx
  module Document
    def self.create(format, opts={})
      if format == :docx
        return WordXML.new(opts)
      else
        raise(UnknownFormatError, format)
      end
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
