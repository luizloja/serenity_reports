module SerenityReport
  class OdsProcessor < BaseProcessor
    def process
      %w(content.xml styles.xml).each do |xml_file|
        next unless @zipfile.find_entry(xml_file)
        evaluate_xml(xml_file)
      end
    end
  end
end
