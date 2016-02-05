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
      if rels_exist?
        @rels = open_xml(doc.zip.read(filename))
      else
        @rels = open_xml(read_fixture(filename))
      end
    end

    def add_honey(url, token)
      #TODO need to escape url??
      self.rels = open_xml(rels.to_xml.gsub('HONEY_TOKEN', "#{url}#{token}"))
      add_file_to_zip(filename, rels.to_xml)
    end

    def add_relation(type, target, target_mode=nil)
      rid = next_rid
      #could be problems if so try childrern[0].add_child
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
