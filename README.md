# Honeydocx

Honeydocx allows you to insert an "invisible" image (1px by 1px) into a document,
when opened the image is fetched the specified url, allowing the user agent and IP
to be collected.

I had plans to extend this to include .pdf, .xlsx and .pptx files and run as an online service. However since all external content is now blocked by default in MS Office 
(An ADS is written to the file by the browser when downloaded) those plans have 
been abandoned.

## Usage
### Create a new document with beacon
```ruby
opts = { url: "http://example.com/image.jpeg", save_path: "beacon.docx" }
Honeydocx::Document.create(:docx, opts)
```

### Add a beacon to an existing document
```ruby
opts = { url: "http://example.com/image.jpeg", path: "file.docx", save_path: "beacon.docx" }
Honeydocx::Document.create(:docx, opts)
```

## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).
