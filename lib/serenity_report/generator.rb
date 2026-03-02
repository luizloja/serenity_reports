require 'fileutils'

module SerenityReport
  module Generator
    def render_odt template_path, output_path = output_name(template_path)
      template = Template.new template_path, output_path
      template.process binding
    end

    private

    def output_name input, output_ext: nil
      ext = File.extname(input)
      base = File.basename(input, ext)
      ext = '.odt' if ext.empty?
      out_ext = output_ext || ext
      tmp_dir = File.expand_path('../../tmp', __dir__)
      FileUtils.mkdir_p(tmp_dir)
      File.join(tmp_dir, "#{base}_output#{out_ext}")
    end
  end
end
