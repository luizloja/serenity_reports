require 'zip'
require 'fileutils'
require 'tmpdir'

module SerenityReport
  class Template
    attr_accessor :template

    def initialize(template, output)
      if output.end_with?('.pdf')
        @pdf_output = output
        source_ext = File.extname(template)
        output = output.sub(/\.pdf$/, source_ext)
      end
      FileUtils.cp(template, output)
      @template = output
    end

    def process(context)
      format = if @template.end_with?('.xlsx')
        :xlsx
      elsif @template.end_with?('.docx')
        :docx
      elsif @template.end_with?('.ods')
        :ods
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
        when :ods
          OdsProcessor.new(zipfile, context, tmpfiles)
        end
        processor.process
      end

      repack_zip(@template)
      convert_to_pdf if @pdf_output
    ensure
      Thread.current[:serenity_report_format] = nil
    end

    private

    def convert_to_pdf
      template_abs = File.expand_path(@template)
      pdf_output_abs = File.expand_path(@pdf_output)
      ext = File.extname(template_abs).delete('.').downcase

      if which('pandoc') && %w[odt docx].include?(ext)
        convert_with_pandoc(template_abs, pdf_output_abs)
      elsif (soffice = find_libreoffice)
        convert_with_libreoffice(soffice, template_abs, pdf_output_abs)
      else
        raise "No PDF converter found. Install pandoc and typst (`brew install pandoc typst`) " \
              "or LibreOffice (`brew install --cask libreoffice`)."
      end

      FileUtils.rm(template_abs)
    end

    def convert_with_pandoc(input, output)
      args = ['pandoc', input, '-o', output]
      args += ['--pdf-engine=typst'] if which('typst')
      success = system(*args)

      unless success && File.exist?(output) && File.size(output) > 0
        raise "PDF conversion with pandoc failed. Ensure a PDF engine is installed (`brew install typst`)."
      end
    end

    def convert_with_libreoffice(soffice, input, output)
      outdir = File.dirname(output)
      FileUtils.mkdir_p(outdir)

      user_installation = Dir.mktmpdir('serenity_lo_')
      begin
        success = system(
          soffice, '--headless', '--norestore',
          "-env:UserInstallation=file://#{user_installation}",
          '--convert-to', 'pdf', '--outdir', outdir, input
        )
        generated = File.join(outdir, File.basename(input, File.extname(input)) + '.pdf')

        unless success && File.exist?(generated) && File.size(generated) > 0
          raise "PDF conversion with LibreOffice failed. " \
                "Ensure LibreOffice is installed correctly and no other instance is running."
        end

        FileUtils.mv(generated, output) unless generated == output
      ensure
        FileUtils.rm_rf(user_installation)
      end
    end

    def find_libreoffice
      candidates = %w[libreoffice soffice]
      candidates.unshift('/Applications/LibreOffice.app/Contents/MacOS/soffice') if RUBY_PLATFORM =~ /darwin/

      candidates.each do |cmd|
        path = cmd.start_with?('/') ? cmd : which(cmd)
        next unless path && File.executable?(path)
        resolved = (File.realpath(path) rescue path)
        next if resolved.include?('OpenOffice')
        return path
      end

      nil
    end

    def which(cmd)
      ENV['PATH'].split(File::PATH_SEPARATOR).each do |dir|
        path = File.join(dir, cmd)
        return path if File.executable?(path)
      end
      nil
    end

    def repack_zip(path)
      tmp_path = "#{path}.tmp"
      Zip::OutputStream.open(tmp_path) do |out|
        Zip::File.open(path) do |zf|
          # ODF spec: mimetype must be first entry, stored uncompressed
          if zf.find_entry('mimetype')
            out.put_next_entry('mimetype', nil, nil, Zip::Entry::STORED)
            out.write(zf.read('mimetype'))
          end
          zf.entries.each do |entry|
            next if entry.directory? || entry.name == 'mimetype'
            out.put_next_entry(entry.name)
            out.write(zf.read(entry.name))
          end
        end
      end
      FileUtils.mv(tmp_path, path)
    end
  end
end
