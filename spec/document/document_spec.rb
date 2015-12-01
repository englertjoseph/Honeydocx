require 'spec_helper'

describe Honeydocx::Document do
  describe '.parse_filetype' do
    it 'identifies a pdf extension' do
      expect(Honeydocx::Document.parse_filetype('test.pdf')).to eq '.pdf'
    end

    it 'identifies a docx extension' do
      expect(Honeydocx::Document.parse_filetype('test.docx')).to eq '.docx'
    end

    it 'identifies an excel extension' do
      expect(Honeydocx::Document.parse_filetype('test.xlsx')).to eq '.xlsx'
    end

    it 'fails for unknown extensions' do
      expect { Honeydocx::Document.parse_filetype('test.xx') }.to raise_error(Honeydocx::Document::UnknownFormatError, "Error the format test.xx is not understood.")
    end
  end
end
