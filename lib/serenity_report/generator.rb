require 'fileutils'

module SerenityReport
  module Generator
    def render_odt template_path, output_path = output_name(template_path)
      template = Template.new template_path, output_path
      template.process binding
    end

    private

    def output_name input
      ext = File.extname(input)
      base = File.basename(input, ext)
      ext = '.odt' if ext.empty?
      tmp_dir = File.expand_path('../../tmp', __dir__)
      FileUtils.mkdir_p(tmp_dir)
      File.join(tmp_dir, "#{base}_output#{ext}")
    end
  end
end
