require 'zip'
require 'nokogiri'

module Honeydocx
  class WordXML

    @@BLANK_PATH = File.expand_path("../word/blank.docx", __FILE__)
    @@HONEY_RELS_HEADER_PATH = File.expand_path("../word/header1.xml.rels", __FILE__)

    attr_reader :path, :zip, :header, :url, :save_path, :token
    attr_accessor :header_rels_xml, :header_xml, :files_to_add, :content_types, :doc_rels, :doc

    def initialize(opts={})
      @path = opts.fetch(:path, WordXML.blank_path)
      @url = opts.fetch(:url)
      @token = opts.fetch(:token)
      @save_path = "#{Dir.pwd}/tmp/#{token}.docx"
      @files_to_add = {}
      open_docx
      add_honey
      save
    end

    def add_honey
      if (new_document?)
        # If using template throw away old header relationships and replace
        @header_rels_xml = template_header_rels
      elsif(has_header?)
        # Add resource && add to header1.xml
        if (!has_header_rels?)
          #create_header_rels
          @header_rels_xml = template_header_rels
          files_to_add['word/_rels/header1.xml.rels'] = @header_rels_xml
          insert_partial
        else
          relations_number = insert_header_rels
          insert_partial(relations_number + 1)
        end
      else
        # Add headers to the document
        # Create header
        @header_xml = File.open(File.expand_path('../word/header1.xml', __FILE__)).read
        files_to_add['word/header1.xml'] = @header_xml
        @header_rels_xml = template_header_rels
        # Add entry to [CONTENT_TYPES].xml
        content_types_patch_file = File.open(File.expand_path('../word/[Content_Types].xml.patch', __FILE__)).read
        content_types_patch = Nokogiri::XML(content_types_patch_file)
        @content_types = Nokogiri::XML(zip.read('[Content_Types].xml'))
        content_types_patch.root.children.each { |child| @content_types.root.add_child(child) }
        # Hack remove the default namespace from the newly added nodes
        @content_types.root.children.each { |child| child.namespace = nil }
        @content_types = @content_types.to_xml

        # Get the header rId
        @doc_rels = Nokogiri::XML(zip.read('word/_rels/document.xml.rels'))
        current_rid = doc_rels.children[0].children.collect {
            |child| child["name"] == "Relationship" }.size + 1
        # TODO are end and foot notes really needed? Leave til later & test
        rids = { 'header1.xml' => current_rid, 'endnotes.xml' => current_rid + 1,
          'footnotes.xml' => current_rid + 2 }
        # Insert into doc_rels
        @doc_rels.root.add_child("<Relationship Id=\"rId#{rids['header1.xml']}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>")
        @doc_rels = @doc_rels.to_xml

        @header_rels_xml = template_header_rels
        files_to_add['word/_rels/header1.xml.rels'] = @header_rels_xml
        #Insert into document.xml
        doc = Nokogiri::XML(zip.read('word/document.xml'))
        doc.at_xpath(".//w:sectPr").prepend_child("<w:headerReference w:type=\"default\" r:id=\"rId#{rids['header1.xml']}\"/>")
        @doc = doc.to_xml
      end
      @header_rels_xml.gsub!("TOKEN_URL", url + token)
    end

    def save
      Zip::OutputStream.open(save_path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)
          if (entry.name == "word/_rels/header1.xml.rels")
            out.write(header_rels_xml)
          elsif (entry.name == "word/header1.xml")
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
              f.puts(data)
            end
          end
        end
      end
    end

    def self.blank_path
      @@BLANK_PATH
    end

    def self.honey_header_path
      @@HONEY_RELS_HEADER_PATH
    end

    private

    def open_docx
      @zip = Zip::File.open(path)
    end

    def template_header_rels
      File.open(WordXML.honey_header_path).read
    end

    def insert_partial(rId = 1)
      # Insert hook into word/header1.xml
      partial = Nokogiri::XML(File.open(File.expand_path("../word/header1.xml.partial" , __FILE__)))
      image_data = partial.at_xpath(".//v:imagedata")
      image_data["r:id"] = "rId#{rId}"
      header_xml =  Nokogiri::XML(zip.read("word/header1.xml"))
      # Add header style if none exist
      if (header_xml.at_xpath(".//w:pPr").nil?)
        partial << Nokogiri::XML(File.open(File.expand_path("../word/pPr.xml", __FILE__)))
      end
      header_xml.at_xpath(".//w:p") << partial.children[0].children.to_xml
      @header_xml = header_xml.to_xml
    end

    def insert_header_rels
      # Get last relationship number (rid)
      # Add relationship with last rid + 1
      # Edit partial to include rid
      @header_rels_xml = zip.read("word/_rels/header1.xml.rels")
      header_rels = Nokogiri::XML(@header_rels_xml)
      relations_number = header_rels.children[0].children.collect {
        |child| child["name"] == "Relationship" }.size
      dict = { "Id" => "rId#{relations_number+1}",
          "Type" => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          "Target" => "TOKEN_URL",
          "TargetMode" => "External"}
      header_rels.children[0].add_child("<Relationship
        Id = \"rId#{relations_number+1}\"
        Type = \"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\"
        Target = \"TOKEN_URL\"
        TargetMode = \"External\"
        />")
      @header_rels_xml = header_rels.to_xml
      relations_number
    end

    def new_document?
      path == WordXML.blank_path
    end

    def has_file?(filename)
      zip.entries.any? { |entry| entry.name == filename }
    end

    def has_header?
      zip.entries.any? { |entry| entry.name == "word/header1.xml" }
    end

    def has_header_rels?
      zip.entries.any? { |entry| entry.name == "word/_rels/header1.xml.rels" }
    end
  end
end
