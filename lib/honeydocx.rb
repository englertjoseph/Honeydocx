require 'pry'
require 'honeydocx/version'
require 'honeydocx/document/document'

module Honeydocx
  def self.create(type, opts={})
    doc = Document.create(type, opts)
  end

  def self.add_honey(filename, opts={})

  end
end
