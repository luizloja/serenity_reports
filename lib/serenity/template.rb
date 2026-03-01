require 'zip'
require 'fileutils'

module Serenity
  class Template
    attr_accessor :template

    def initialize(template, output)
      FileUtils.cp(template, output)
      @template = output
    end

    def process(context)
      format = @template.end_with?('.docx') ? :docx : :odt
      tmpfiles = []
      Thread.current[:serenity_format] = format

      Zip::File.open(@template) do |zipfile|
        processor = if format == :odt
          OdtProcessor.new(zipfile, context, tmpfiles)
        else
          DocxProcessor.new(zipfile, context, tmpfiles)
        end
        processor.process
      end
    ensure
      Thread.current[:serenity_format] = nil
    end
  end
end
