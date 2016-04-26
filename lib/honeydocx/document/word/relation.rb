require_relative 'word_helper'

module Honeydocx
  class Relation
    include WordHelper

    attr_reader :filename, :doc
    attr_accessor :rels
    #Initially only for header rels
    def initialize(filename, doc)
      @filename = filename
      @doc = doc
      rels_file = rels_exist? ? doc.zip.read(filename) : read_fixture(filename)
      @rels = open_xml(rels_file)
    end

    def add_honey(url)
      self.rels = open_xml(rels.to_xml.gsub('HONEY_TOKEN', url))
      add_file_to_zip(filename, rels.to_xml)
    end

    def add_relation(type, target, target_mode=nil)
      rid = next_rid
      rels.root.add_child("<Relationship Id=\"rId#{rid}\" Type=\"#{type}\" Target=\"#{target}\" #{"TargetMode=\"#{target_mode}\"" if target_mode }/>")
      rid
    end

    def to_xml
      @rels.to_xml
    end

    private

    def template_header_rels
      File.open(WordXML.honey_header_rels_path).read
    end

    def rels_exist?
      #Name for header = word/_rels/header1.xml.rels
      doc.zip.entries.any? { |entry| entry.name == filename }
    end

    def next_rid
      rels.children[0].children.collect {
            |child| child["name"] == "Relationship" }.size + 1
    end
  end
end
