require 'zip'
require 'nokogiri'
require_relative 'word/header'
require_relative 'word/word_helper'

module Honeydocx
  class WordXML
    include WordHelper

    @@Fixtures_path = File.expand_path("../word_fixtures" , __FILE__)
    @@Blank_path = File.expand_path('blank_template.docx', @@Fixtures_path)
    @@Honey_rels_header_path = File.expand_path('header1.xml.rels', @@Fixtures_path)

    attr_reader :path, :zip, :header, :url, :files_to_add

    def initialize(opts={})
      @path = opts.fetch(:path, WordXML.blank_path)
      @url = opts.fetch(:url)
      @files_to_add = {}
      @zip = open_docx
      @header = Header.new(self)
      add_honey(url)
    end

    def add_honey(url)
      header.add_honey(url)
    end

    def save(save_path)
      # Create new word file
      Zip::OutputStream.open(save_path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)
          if (entry.file?)
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
      files_to_add.each do |filename, data|
        Zip::File.open(save_path, Zip::File::CREATE) do |zipfile|
          zipfile.get_output_stream(filename) { |f| f.puts(data) }
        end
      end
      save_path
    end

    def add_file_to_zip(filename, data)
      self.files_to_add[filename] = data
    end

    def self.blank_path
      @@Blank_path
    end

    def new_document?
        path == WordXML.blank_path
    end

    private

      def open_docx
        Zip::File.open(path)
      end

      def has_file?(filename)
        zip.entries.any? { |entry| entry.name == filename }
      end
  end
end
