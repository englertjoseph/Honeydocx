require_relative 'word_helper'
require_relative 'relation'

module Honeydocx
  class Header
    include WordHelper

    attr_reader :doc, :header_rels, :doc_rid, :header_rid
    attr_accessor :header

    def initialize(doc)
      @doc = doc
      @header = open_xml(doc.zip.read('word/header1.xml')) if header_exists?
      @header_rels = Relation.new('word/_rels/header1.xml.rels', doc)
    end

    def add_honey(url, token)
      # if header_exits?
      #   @header = open_xml(doc.zip.read("word/header1.xml"))
      #   edit_header
      # else
      #   add_header
      # end
      if (doc.new_document?)
        header_rels.add_honey(url, token)
      else
        add_relations(url, token)
        edit_header
        patch_content_types
        add_file_to_zip('word/header1.xml', header.to_xml)
      end
    end

    private
      def header_exists?
        !!doc.zip.find_entry('word/header1.xml')
      end

      def edit_header
        if !header_exists?
          add_header
        end
        insert_partial(header_rid)
      end

      def add_header
        @header = open_xml(read_fixture('word/header1.xml'))
        add_relation_to_document
        reference_header_in_document
      end

      def add_relations(url, token)
        insert_into_header_rels(url+token)
      end

      def patch_content_types
          # Add entry to [CONTENT_TYPES].xml
          content_types_patch_file = read_fixture('[Content_Types].xml.patch')
          content_types_patch = Nokogiri::XML(content_types_patch_file)
          content_types = Nokogiri::XML(doc.zip.read('[Content_Types].xml'))
          content_types_patch.root.children.each { |child| content_types.root.add_child(child) }
          # Hack remove the default namespace from the newly added nodes
          content_types.root.children.each { |child| child.namespace = nil }
          content_types = content_types.to_xml
          add_file_to_zip('[Content_Types].xml', content_types)
        end

      def insert_partial(rId = 1)
        # Insert hook into word/header1.xml
        partial = Nokogiri::XML(read_fixture('header1.xml.partial'))
        image_data = partial.at_xpath(".//v:imagedata")
        image_data["r:id"] = "rId#{rId}"
        header_xml =  header#Nokogiri::XML(doc.zip.read("word/header1.xml"))
        # Add header style if none exist
        if (header_xml.at_xpath(".//w:pPr").nil?)
          partial << Nokogiri::XML(read_fixture('pPr.xml'))
        end
        header_xml.at_xpath(".//w:p") << partial.children[0].children.to_xml
        self.header = header_xml
      end

      def add_relation_to_document
        doc_rels = Relation.new('word/_rels/document.xml.rels', doc)
        @doc_rid = doc_rels.add_relation('http://schemas.openxmlformats.org/officeDocument/2006/relationships/header', 'header1.xml')
        add_file_to_zip('word/_rels/document.xml.rels', doc_rels.to_xml)
      end

      def reference_header_in_document
        document = Nokogiri::XML(doc.zip.read('word/document.xml'))
        document.at_xpath(".//w:sectPr").prepend_child("<w:headerReference w:type=\"default\" r:id=\"rId#{doc_rid}\"/>")
        add_file_to_zip('word/document.xml', document.to_xml)
      end

      def insert_into_header_rels(token)
        # Get last relationship number (rid)
        # Add relationship with last rid + 1
        # Edit partial to include rid
        @header_rid = header_rels.add_relation('http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', token, "External")
        add_file_to_zip('word/_rels/header1.xml.rels', header_rels.to_xml)
      end
  end
end
