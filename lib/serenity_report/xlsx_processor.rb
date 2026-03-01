require 'nokogiri'

module SerenityReport
  class XlsxProcessor < BaseProcessor
    TEMPLATE_PATTERN = /\{%.*?%\}/m
    NOKOGIRI_SAVE_OPTS = Nokogiri::XML::Node::SaveOptions::AS_XML | Nokogiri::XML::Node::SaveOptions::NO_DECLARATION

    def process
      strings = load_shared_strings
      indices = find_template_indices(strings)
      return if indices.empty?

      # Build old→new index map for non-template shared strings
      index_map = {}
      strings.each_with_index do |s, i|
        next if indices.include?(i)
        index_map[i] = index_map.size
      end

      # Collect all output strings across sheets for the final shared strings table
      all_output_strings = []

      find_sheet_files.each do |sheet_file|
        sheet_strings = process_sheet(sheet_file, strings, indices, index_map)
        all_output_strings.concat(sheet_strings)
      end

      # Build final shared strings: non-template originals + new output strings
      final_strings = []
      strings.each_with_index do |s, i|
        final_strings << s unless indices.include?(i)
      end
      final_strings.concat(all_output_strings)

      update_shared_strings(final_strings)
    end

    private

    def load_shared_strings
      xml = @zipfile.read('xl/sharedStrings.xml')
      doc = Nokogiri::XML(xml)
      doc.remove_namespaces!
      doc.xpath('//si/t').map(&:text)
    end

    def find_template_indices(strings)
      indices = Set.new
      strings.each_with_index do |s, i|
        indices.add(i) if s.match?(TEMPLATE_PATTERN)
      end
      indices
    end

    def find_sheet_files
      @zipfile.entries.map(&:name).select { |name| name.match?(%r{\Axl/worksheets/sheet\d+\.xml\z}) }
    end

    def inline_shared_strings(xml, strings, indices, index_map)
      doc = Nokogiri::XML(xml)
      ns = doc.root.namespace&.href

      xpath = ns ? '//xmlns:c[@t="s"]' : '//c[@t="s"]'
      doc.xpath(xpath).each do |cell|
        v_node = cell.at_xpath(ns ? 'xmlns:v' : 'v')
        next unless v_node

        idx = v_node.text.to_i

        if indices.include?(idx)
          cell['t'] = 'inlineStr'
          v_node.remove

          is_node = Nokogiri::XML::Node.new('is', doc)
          t_node = Nokogiri::XML::Node.new('t', doc)
          t_node.content = strings[idx]
          is_node.add_child(t_node)
          cell.add_child(is_node)
        elsif index_map.key?(idx)
          v_node.content = index_map[idx].to_s
        end
      end

      xml_declaration(doc) + doc.root.to_xml(save_with: NOKOGIRI_SAVE_OPTS)
    end

    # Converts inlineStr cells back to shared string references.
    # Returns the array of new string values collected from the sheet.
    def convert_inline_to_shared(xml, base_index)
      doc = Nokogiri::XML(xml)
      ns = doc.root.namespace&.href

      new_strings = []
      xpath = ns ? '//xmlns:c[@t="inlineStr"]' : '//c[@t="inlineStr"]'
      doc.xpath(xpath).each do |cell|
        is_node = cell.at_xpath(ns ? 'xmlns:is' : 'is')
        next unless is_node

        t_node = is_node.at_xpath(ns ? 'xmlns:t' : 't')
        text = t_node ? t_node.text : ''

        is_node.remove
        cell['t'] = 's'

        v_node = Nokogiri::XML::Node.new('v', doc)
        v_node.content = (base_index + new_strings.size).to_s
        cell.add_child(v_node)

        new_strings << text
      end

      [xml_declaration(doc) + doc.root.to_xml(save_with: NOKOGIRI_SAVE_OPTS), new_strings]
    end

    def process_sheet(file, strings, indices, index_map)
      xml = @zipfile.read(file)
      xml = inline_shared_strings(xml, strings, indices, index_map)

      odteruby = OdtEruby.new(XmlReader.new(xml))
      out = odteruby.evaluate(@context)
      out.force_encoding Encoding.default_external

      out = cleanup_rows(out)

      # Count of non-template shared strings (they occupy indices 0..n-1)
      base_index = index_map.size
      out, new_strings = convert_inline_to_shared(out, base_index)

      @tmpfiles << (tmpfile = Tempfile.new("serenity_report"))
      tmpfile << out
      tmpfile.close
      @zipfile.replace(file, tmpfile.path)

      new_strings
    end

    def update_shared_strings(final_strings)
      doc = Nokogiri::XML(@zipfile.read('xl/sharedStrings.xml'))
      ns = doc.root.namespace&.href

      sst = doc.at_xpath(ns ? '//xmlns:sst' : '//sst')
      sst.children.remove

      sst['count'] = final_strings.size.to_s
      sst['uniqueCount'] = final_strings.size.to_s

      final_strings.each do |str|
        si = Nokogiri::XML::Node.new('si', doc)
        si.namespace = sst.namespace
        t = Nokogiri::XML::Node.new('t', doc)
        t.namespace = sst.namespace
        t.content = str
        si.add_child(t)
        sst.add_child(si)
      end

      @tmpfiles << (file = Tempfile.new("serenity_report"))
      file << xml_declaration(doc) + doc.root.to_xml(save_with: NOKOGIRI_SAVE_OPTS)
      file.close
      @zipfile.replace('xl/sharedStrings.xml', file.path)
    end

    def cleanup_rows(xml)
      doc = Nokogiri::XML(xml)
      ns = doc.root.namespace&.href

      sheet_data = doc.at_xpath(ns ? '//xmlns:sheetData' : '//sheetData')
      return xml unless sheet_data

      rows = sheet_data.xpath(ns ? 'xmlns:row' : 'row')

      rows.each do |row|
        cells = row.xpath(ns ? 'xmlns:c' : 'c')
        row.remove if cells.empty?
      end

      remaining_rows = sheet_data.xpath(ns ? 'xmlns:row' : 'row')
      remaining_rows.each_with_index do |row, i|
        new_row_num = i + 1
        row['r'] = new_row_num.to_s

        cells = row.xpath(ns ? 'xmlns:c' : 'c')
        cells.each do |cell|
          ref = cell['r']
          next unless ref
          col_letters = ref.gsub(/\d+/, '')
          cell['r'] = "#{col_letters}#{new_row_num}"
        end
      end

      xml_declaration(doc) + doc.root.to_xml(save_with: NOKOGIRI_SAVE_OPTS)
    end

    def xml_declaration(_doc)
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    end
  end
end
