require 'zip'
require 'fileutils'

module SerenityReport
  class Template
    attr_accessor :template

    def initialize(template, output)
      FileUtils.cp(template, output)
      @template = output
    end

    def process(context)
      format = if @template.end_with?('.xlsx')
        :xlsx
      elsif @template.end_with?('.docx')
        :docx
      else
        :odt
      end
      tmpfiles = []
      Thread.current[:serenity_report_format] = format

      Zip::File.open(@template) do |zipfile|
        processor = case format
        when :odt
          OdtProcessor.new(zipfile, context, tmpfiles)
        when :docx
          DocxProcessor.new(zipfile, context, tmpfiles)
        when :xlsx
          XlsxProcessor.new(zipfile, context, tmpfiles)
        end
        processor.process
      end

      repack_zip(@template) if format == :xlsx
    ensure
      Thread.current[:serenity_report_format] = nil
    end

    private

    def repack_zip(path)
      tmp_path = "#{path}.tmp"
      Zip::OutputStream.open(tmp_path) do |out|
        Zip::File.open(path) do |zf|
          zf.entries.each do |entry|
            out.put_next_entry(entry.name)
            out.write(zf.read(entry.name))
          end
        end
      end
      FileUtils.mv(tmp_path, path)
    end
  end
end
