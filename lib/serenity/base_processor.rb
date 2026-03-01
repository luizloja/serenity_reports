require 'tempfile'

module Serenity
  class BaseProcessor
    def initialize(zipfile, context, tmpfiles)
      @zipfile = zipfile
      @context = context
      @tmpfiles = tmpfiles
    end

    def process
      raise NotImplementedError
    end

    private

    def evaluate_xml(xml_file)
      content = @zipfile.read(xml_file)
      yield content if block_given?

      odteruby = OdtEruby.new(XmlReader.new(content))
      out = odteruby.evaluate(@context)
      out.force_encoding Encoding.default_external

      @tmpfiles << (file = Tempfile.new("serenity"))
      file << out
      file.close
      @zipfile.replace(xml_file, file.path)
      out
    end
  end
end
