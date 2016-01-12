require 'zip'
require 'nokogiri'
require_relative 'word/header'
require_relative 'word/word_helper'

module Honeydocx
  class WordXML
    include WordHelper

    @@Fixtures_path = File.expand_path("../word_fixtures" , __FILE__)
    @@Blank_path = File.expand_path('blank.docx', @@Fixtures_path)
    @@Honey_rels_header_path = File.expand_path('header1.xml.rels', @@Fixtures_path)

    attr_reader :path, :zip, :header, :url, :save_path, :token
    attr_accessor :header_rels_xml, :header_xml, :files_to_add, :content_types, :doc_rels, :doc

    def initialize(opts={})
      @path = opts.fetch(:path, WordXML.blank_path)
      @url = opts.fetch(:url)
      @token = opts.fetch(:token)
      @save_path = "#{Dir.pwd}/tmp/#{token}.docx"
      @files_to_add = {}
      open_docx
      @header = Header.new(self)

      add_honey(url, token)
      save
    end

    def add_honey(url, token)
      header.add_honey(url, token)
      patch_content_types
    end

    def save
      Zip::OutputStream.open(save_path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)
          if (entry.name == "word/_rels/header1.xml.rels" && header_rels_xml)
            out.write(header_rels_xml)
          elsif (entry.name == "word/header1.xml" && header_xml)
            out.write(header_xml)
          elsif (entry.name == "[Content_Types].xml" && content_types)
            out.write(content_types)
          elsif (entry.name == "word/_rels/document.xml.rels" && doc_rels)
            out.write(doc_rels)
          elsif (entry.name == "word/document.xml" && doc)
            out.write(doc)
          elsif (entry.file?)
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
      files_to_add.each do |filename, data|
        if (!has_file?(filename))
          Zip::File.open(save_path, Zip::File::CREATE) do |zipfile|
            zipfile.get_output_stream(filename) do |f|
              ##TODO change so this is the only save.. overwite files if they
              #exist perhaps by setting f to ""?
              f.puts(data)
            end
          end
        end
      end
    end

    def add_file_to_zip(filename, data)
      files_to_add[filename] = data
    end

    def self.blank_path
      @@Blank_path
    end

    def self.honey_header_rels_path
      @@Honey_rels_header_path
    end

    private

      def open_docx
        @zip = Zip::File.open(path)
      end

      def patch_content_types
          # Add entry to [CONTENT_TYPES].xml
          content_types_patch_file = File.open(File.expand_path('../word_fixtures/[Content_Types].xml.patch', __FILE__)).read
          content_types_patch = Nokogiri::XML(content_types_patch_file)
          content_types = Nokogiri::XML(zip.read('[Content_Types].xml'))
          content_types_patch.root.children.each { |child| content_types.root.add_child(child) }
          # Hack remove the default namespace from the newly added nodes
          content_types.root.children.each { |child| child.namespace = nil }
          content_types = content_types.to_xml
          add_file_to_zip('[Content_Types].xml', content_types)
        end

        def has_file?(filename)
          zip.entries.any? { |entry| entry.name == filename }
        end
  end
end
