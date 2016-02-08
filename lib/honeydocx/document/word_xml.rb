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
      open_docx
      @header = Header.new(self)
      add_honey(url)
    end

    def add_honey(url)
      header.add_honey(url)
    end

    def save
      temp_file = Tempfile.new(['test', '.docx'])
      Zip::OutputStream.open(temp_file) { |zos| }
      Zip::OutputStream.open(temp_file) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)
          if (entry.file?)
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
      self.files_to_add.each do |filename, data|
        Zip::File.open(temp_file, Zip::File::CREATE) do |zipfile|
          zipfile.get_output_stream(filename) do |f|
            ##TODO change so this is the only save.. overwite files if they
            #exist perhaps by setting f to ""?
            f.puts(data)
          end
        end
      end
      temp_file
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
        @zip = Zip::File.open(path)
      end

      def has_file?(filename)
        zip.entries.any? { |entry| entry.name == filename }
      end
  end
end
