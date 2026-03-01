module Serenity
  class DocxProcessor < BaseProcessor
    def process
      xml_files = ['word/document.xml']

      # Also process headers and footers if present
      @zipfile.entries.each do |entry|
        if entry.name.match?(%r{\Aword/(header|footer)\d*\.xml\z})
          xml_files << entry.name
        end
      end

      xml_files.each do |xml_file|
        next unless @zipfile.find_entry(xml_file)
        evaluate_xml(xml_file)
      end
    end
  end
end
