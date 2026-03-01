module Serenity
  module Generator
    def render_odt template_path, output_path = output_name(template_path)
      template = Template.new template_path, output_path
      template.process binding
    end

    private

    def output_name input
      ext = File.extname(input)
      base = input.chomp(ext)
      ext = '.odt' if ext.empty?
      "#{base}_output#{ext}"
    end
  end
end
