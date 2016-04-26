require 'honeydocx/version'
require 'honeydocx/document/document'

module Honeydocx
  def self.create(type, opts={})
    doc = Document.create(type, opts)
  end
end
