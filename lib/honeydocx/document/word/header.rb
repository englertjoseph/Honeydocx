require_relative 'word_helper'
require_relative 'relation'

module Honeydocx
  class Header
    include WordHelper

    attr_reader :doc, :header_rels, :doc_rid, :header_rid

    def initialize(doc)
      @doc = doc
    end

    def add_honey(url, token)
      if header_exits?
        edit_header
      else
        add_header
      end
    end

    private
      def header_exits?
        !!doc.zip.find_entry('word/header1.xml')
      end

      def edit_header
        # Add resource && add to header1.xml
        if (!has_header_rels?)
          add_header_relation('word/_rels/header1.xml.rels')
          insert_partial
        else
          insert_into_header_rels
          fail if header_rid.nil?
          insert_partial(header_rid)
        end
      end

      def add_header
        add_header_relation('word/_rels/header1.xml.rels')
        # Add headers to the document
        # Create header
        header_xml = read_fixture('header1.xml')
        add_file_to_zip('word/header1.xml', header_xml)
        add_relation_to_document
        reference_header_in_document
      end

      def add_header_relation(filename)
        @header_rels = Relation.new(filename, doc)
        #add relation with id... therefore need to reomve refernce from fixtures
        insert_into_header_rels
      end

      def insert_partial(rId = 1)
        # Insert hook into word/header1.xml
        partial = Nokogiri::XML(read_fixture('header1.xml.partial'))
        image_data = partial.at_xpath(".//v:imagedata")
        image_data["r:id"] = "rId#{rId}"
        header_xml =  Nokogiri::XML(doc.zip.read("word/header1.xml"))
        # Add header style if none exist
        if (header_xml.at_xpath(".//w:pPr").nil?)
          partial << Nokogiri::XML(read_fixture('pPr.xml'))
        end
        header_xml.at_xpath(".//w:p") << partial.children[0].children.to_xml
        @header_xml = header_xml.to_xml
      end

      def add_relation_to_document
        doc_rels = Relation.new('word/_rels/document.xml.rels', doc)
        @doc_rid = doc_rels.add_relation('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml')
        add_file_to_zip('word/_rels/document.xml.rels', doc_rels.to_xml)
      end

      def reference_header_in_document
        document = Nokogiri::XML(doc.zip.read('word/document.xml'))
        document.at_xpath(".//w:sectPr").prepend_child("<w:headerReference w:type=\"default\" r:id=\"rId#{doc_rid}\"/>")
        add_file_to_zip('word/document.xml', document)
      end

      def insert_into_header_rels
        # Get last relationship number (rid)
        # Add relationship with last rid + 1
        # Edit partial to include rid
        @header_rels = Relation.new('word/_rels/header1.xml.rels', doc)
        @header_rid = header_rels.add_relation('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', 'TOKEN_URL')
        add_file_to_zip('word/_rels/header1.xml.rels', header_rels.to_xml)
      end
  end
end
