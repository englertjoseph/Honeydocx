require 'spec_helper'

describe Honeydocx::WordXML do
  let(:dummy_document) { Class.new { extend Honeydocx::Document } }
  describe '#new' do
    before(:each) do
      @url = "http://localhost/"
      @token = "dog.jpg"
    end

    context 'using template document' do
      before(:each) do
        @opts = { url: @url, token: @token }
      end

      it 'should set path to template document if no path specified' do
        wordXML = dummy_document.create(:docx, @opts)
        expect(wordXML.path).to eq(Honeydocx::WordXML.blank_path)
      end

      it 'should insert a token in the header rels' do
        wordXML = dummy_document.create(:docx, @opts)
        expected_xml = File.open(File.expand_path('../../fixtures/header1.xml.rels', __FILE__)).read
        expect(clean_xml(wordXML.header_rels_xml)).to eq (clean_xml(expected_xml))
      end
    end

    context 'Using supplied document with header' do
      context 'with no header rels' do
        before(:each) do
          @path = File.expand_path('../../fixtures/header_no_rels.docx', __FILE__)
          @opts = { path: @path, url: @url, token: @token }
        end

        it 'should set the path of the document' do
          wordXML = dummy_document.create(:docx, @opts)
          expect(wordXML.path).to eq(@opts[:path])
        end

        it 'should create header rels with token' do
          wordXML = dummy_document.create(:docx, @opts)
          expected_xml = File.open(File.expand_path('../../fixtures/header1.xml.rels', __FILE__)).read
          expect(clean_xml(wordXML.header_rels_xml)).to eq (clean_xml(expected_xml))
        end

        it 'should create a header rels file' do
          wordXML = dummy_document.create(:docx, @opts)
          expect(wordXML.send(:has_header_rels?)).to eq(true)
        end

        it 'should reference token in word/header1.xml' do
          wordXML = dummy_document.create(:docx, @opts)
          expect(clean_xml(wordXML.header_xml)).to eq (clean_xml(expected_header))
        end
      end

      context 'with header rels' do
        before(:each) do
          @path = File.expand_path('../../fixtures/header_with_rels.docx', __FILE__)
          @opts = { path: @path, url: @url, token: @token }
        end

        it 'should set the path of the document' do
          wordXML = dummy_document.create(:docx, @opts)
          expect(wordXML.path).to eq(@opts[:path])
        end

        it 'should insert a token in the header rels' do
          wordXML = dummy_document.create(:docx, @opts)
          expected_xml = File.open(File.expand_path('../../fixtures/header1.xml.rels', __FILE__)).read
          expect(clean_xml(wordXML.header_rels_xml)).to eq (clean_xml(expected_xml))
        end
      end
    end

    context 'Using supplied document with no header' do
      before(:each) do

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
