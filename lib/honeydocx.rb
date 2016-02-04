require 'pry'
require 'honeydocx/version'
require 'honeydocx/document/document'

module Honeydocx
  def self.create(type, opts={})
    doc = Document.create(type, opts)
    doc.save!
  end

  def self.add_honey(filename, opts={})

  end
end
