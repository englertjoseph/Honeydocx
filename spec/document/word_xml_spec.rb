require 'spec_helper'

describe Honeydocx::WordXML do
  let(:dummy_document) { Class.new { extend Honeydocx::Document } }
  describe '#new' do
    before(:each) do
      @url = "http://localhost/"
      @token = "dog.jpg"
      @save_path = "#{Dir.pwd}/tmp/#{@token}.docx" #Dirty fix allow a doc to be created with a save path
    end

    context 'using template document' do
      before(:each) do
        @opts = { url: @url, token: @token }
        @wordXML = dummy_document.create(:docx, @opts)
      end

      it 'should set path to template document if no path specified' do
        expect(@wordXML.path).to eq(Honeydocx::WordXML.blank_path)
      end

      it 'should insert a token in the header rels' do
        expected_xml = File.open(File.expand_path('../../fixtures/header1.xml.rels', __FILE__)).read
        expect(clean_xml(@wordXML.header_rels_xml)).to eq (clean_xml(expected_xml))
      end
    end

    context 'Using supplied document with header' do
      context 'with no header rels' do
        before(:each) do
          @path = File.expand_path('../../fixtures/header_no_rels.docx', __FILE__)
          @opts = { path: @path, url: @url, token: @token }
          @wordXML = dummy_document.create(:docx, @opts)
        end

        it 'should set the path of the document' do
          expect(@wordXML.path).to eq(@opts[:path])
        end

        it 'should create header rels with token' do
          expected_xml = File.open(File.expand_path('../../fixtures/header1.xml.rels', __FILE__)).read
          expect(clean_xml(@wordXML.header_rels_xml)).to eq (clean_xml(expected_xml))
        end

        it 'should create a header rels file' do
          header_rels = get_header_rels(@save_path)
          expect(clean_xml(header_rels)).to eq(clean_xml(expected_header_rels))
        end

        it 'should reference token in word/header1.xml', focus: true do
          expect(clean_xml(@wordXML.header_xml)).to eq (clean_xml(expected_header))
        end
      end

      context 'with header rels' do
        before(:each) do
          @path = File.expand_path('../../fixtures/header_with_rels.docx', __FILE__)
          @opts = { path: @path, url: @url, token: @token }
          @wordXML = dummy_document.create(:docx, @opts)
        end

        it 'should set the path of the document' do
          expect(@wordXML.path).to eq(@opts[:path])
        end

        it 'should insert a token in the header rels' do
          expected_xml = File.open(File.expand_path('../../fixtures/header1_with_rels.xml.rels', __FILE__)).read
          expect(clean_xml(@wordXML.header_rels_xml)).to eq (clean_xml(expected_xml))
        end

        it 'should retain the old rels' do
          old_rels = Nokogiri::XML(get_header_rels(@path))
          old_rels.remove_namespaces! # XML namespaces aren't correct and no searchable without removing
          old_rels.root.xpath(".//Relationship").each do |relation|
            expect(@wordXML.header_rels_xml).to include(relation)
          end
        end
      end
    end

    context 'Using supplied document with no header' do
      before(:each) do
        @path = File.expand_path('../../fixtures/blank_no_header.docx', __FILE__)
        @opts = { path: @path, url: @url, token: @token }
        @wordXML = dummy_document.create(:docx, @opts)
      end

      it 'should set the path of the document' do
        expect(@wordXML.path).to eq(@opts[:path])
      end

      it 'should create a header1.xml file' do
        header = get_header(@save_path)
        expect(clean_xml(header)).to eq(clean_xml(expected_header))
      end

      it 'should register header in [CONTENT_TYPES].xml' do
        header = "<Override PartName=\"/word/header1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml\"/>"
        expect(open_xml('[CONTENT_TYPES].xml', @save_path)).to include(header)
      end
    end

    it 'should throw exception if url not specified' do
      expect { dummy_document.create(:docx, { token: @token }) }.to raise_error(KeyError)
    end

    it 'should throw exception if token not specified' do
      expect { dummy_document.create(:docx, { url: @url }) }.to raise_error(KeyError)
    end
  end
end
