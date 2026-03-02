# encoding: utf-8
require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

module SerenityReport
  describe OdtProcessor do
    it "processes a document with simple variable substitution" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('odt/variables.odt'), tmp('output_variables.odt'))
      template.process binding

      expect(tmp('output_variables.odt')).to contain_in('content.xml', 'Malcolm Reynolds')
      expect(tmp('output_variables.odt')).to contain_in('content.xml', 'captain')
    end

    it "unrolls a simple for loop" do
      @crew = %w{'River', 'Jayne', 'Wash'}

      template = Template.new(fixture('odt/loop.odt'), tmp('output_loop.odt'))
      template.process binding
    end

    it "unrolls an advanced loop with tables" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('odt/loop_table.odt'), tmp('output_loop_table.odt'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_loop_table.odt')).to contain_in('content.xml', text)
      end
    end

    it "processes an advanced document" do
      @persons = [
        Person.new('Malcolm', 'captain',    10.5, 20.3, 30.1),
        Person.new('River',   'psychic',    40.2, 50.7, 60.4),
        Person.new('Jay',     'gunslinger', 70.8, 80.9, 90.6)
      ]

      template = Template.new(fixture('odt/advanced.odt'), tmp('output_advanced.odt'))
      template.process binding

      ['Malcolm', 'captain', 'River', 'psychic', 'Jay', 'gunslinger'].each do |text|
        expect(tmp('output_advanced.odt')).to contain_in('content.xml', text)
      end

      { 'Object 1' => 'Malcolm', 'Object 2' => 'River', 'Object 3' => 'Jay' }.each do |obj, name|
        expect(tmp('output_advanced.odt')).to contain_in("#{obj}/content.xml", name)
      end

      expect(tmp('output_advanced.odt')).not_to contain_in('content.xml', '[*]')
    end

    it "processes a greek document" do
      @h = {'ελληνικο' => 'κειμενο'}
      template = Template.new(fixture('odt/greek.odt'), tmp('output_greek.odt'))
      template.process binding
      expect(tmp('output_greek.odt')).to contain_in('content.xml', 'κειμενο')
    end

    it "loops and generates table rows" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('odt/table_rows.odt'), tmp('output_table_rows.odt'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_table_rows.odt')).to contain_in('content.xml', text)
      end
    end

    it "parses the header" do
      @title = 'captain'

      template = Template.new(fixture('odt/header.odt'), tmp('output_header.odt'))
      template.process(binding)
      expect(tmp('output_header.odt')).to contain_in('styles.xml', 'captain')
    end

    it 'parses the footer' do
      @title = 'captain'

      template = Template.new(fixture('odt/footer.odt'), tmp('output_footer.odt'))
      template.process(binding)
      expect(tmp('output_footer.odt')).to contain_in('styles.xml', 'captain')
    end

    it "deduplicates table names when tables are inside loops" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('odt/loop_table.odt'), tmp('output_dedup_tables.odt'))
      template.process binding

      content = Zip::File.open(tmp('output_dedup_tables.odt')) { |z| z.read('content.xml') }
      table_names = content.scan(/table:name="([^"]+)"/).flatten
      expect(table_names).to eq(table_names.uniq), "Expected unique table names but found duplicates: #{table_names}"
    end

    it "produces a ZIP with no data descriptor flags" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('odt/variables.odt'), tmp('output_zip_flags.odt'))
      template.process binding

      Zip::File.open(tmp('output_zip_flags.odt')) do |zf|
        zf.entries.each do |entry|
          flag = entry.gp_flags & 0x0008
          expect(flag).to eq(0), "Entry '#{entry.name}' has data descriptor flag set (gp_flags=0x#{entry.gp_flags.to_s(16)})"
        end
      end
    end

    it "stores mimetype as first entry uncompressed in ODT" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('odt/variables.odt'), tmp('output_mimetype.odt'))
      template.process binding

      Zip::File.open(tmp('output_mimetype.odt')) do |zf|
        first_entry = zf.entries.first
        expect(first_entry.name).to eq('mimetype')
        expect(first_entry.compression_method).to eq(Zip::Entry::STORED)
      end
    end
  end

  describe DocxProcessor do
    it "processes a document with simple variable substitution" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('docx/variables.docx'), tmp('output_variables.docx'))
      template.process binding

      expect(tmp('output_variables.docx')).to contain_in('word/document.xml', 'Malcolm Reynolds')
      expect(tmp('output_variables.docx')).to contain_in('word/document.xml', 'captain')
    end

    it "unrolls a simple for loop" do
      @crew = %w{'River', 'Jayne', 'Wash'}

      template = Template.new(fixture('docx/loop.docx'), tmp('output_loop.docx'))
      template.process binding
    end

    it "unrolls an advanced loop with tables" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('docx/loop_table.docx'), tmp('output_loop_table.docx'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_loop_table.docx')).to contain_in('word/document.xml', text)
      end
    end

    it "processes an advanced document" do
      @persons = [
        Person.new('Malcolm', 'captain',    10.5, 20.3, 30.1),
        Person.new('River',   'psychic',    40.2, 50.7, 60.4),
        Person.new('Jay',     'gunslinger', 70.8, 80.9, 90.6)
      ]

      template = Template.new(fixture('docx/advanced.docx'), tmp('output_advanced.docx'))
      template.process binding

      ['Malcolm', 'captain', 'River', 'psychic', 'Jay', 'gunslinger'].each do |text|
        expect(tmp('output_advanced.docx')).to contain_in('word/document.xml', text)
      end
    end

    it "processes a greek document" do
      @h = {'ελληνικο' => 'κειμενο'}
      template = Template.new(fixture('docx/greek.docx'), tmp('output_greek.docx'))
      template.process binding
      expect(tmp('output_greek.docx')).to contain_in('word/document.xml', 'κειμενο')
    end

    it "loops and generates table rows" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('docx/table_rows.docx'), tmp('output_table_rows.docx'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_table_rows.docx')).to contain_in('word/document.xml', text)
      end
    end

    it "parses the header" do
      @title = 'captain'

      template = Template.new(fixture('docx/header.docx'), tmp('output_header.docx'))
      template.process(binding)
      expect(tmp('output_header.docx')).to contain_in('word/header1.xml', 'captain')
    end

    it 'parses the footer' do
      @title = 'captain'

      template = Template.new(fixture('docx/footer.docx'), tmp('output_footer.docx'))
      template.process(binding)
      expect(tmp('output_footer.docx')).to contain_in('word/footer1.xml', 'captain')
    end
  end

  describe OdsProcessor do
    it "processes a document with simple variable substitution" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('ods/variables.ods'), tmp('output_variables.ods'))
      template.process binding

      expect(tmp('output_variables.ods')).to contain_in('content.xml', 'Malcolm Reynolds')
      expect(tmp('output_variables.ods')).to contain_in('content.xml', 'captain')
    end

    it "unrolls a simple for loop" do
      @crew = %w{'River', 'Jayne', 'Wash'}

      template = Template.new(fixture('ods/loop.ods'), tmp('output_loop.ods'))
      template.process binding
    end

    it "unrolls an advanced loop with tables" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('ods/loop_table.ods'), tmp('output_loop_table.ods'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_loop_table.ods')).to contain_in('content.xml', text)
      end
    end

    it "processes an advanced document" do
      PersonOds = Struct.new(:nome) unless defined?(PersonOds)
      @persons = [
        PersonOds.new('Malcolm'),
        PersonOds.new('River'),
        PersonOds.new('Jay')
      ]

      template = Template.new(fixture('ods/advanced.ods'), tmp('output_advanced.ods'))
      template.process binding

      ['Malcolm', 'River', 'Jay'].each do |text|
        expect(tmp('output_advanced.ods')).to contain_in('content.xml', text)
      end
    end

    it "processes a greek document" do
      @h = {'ελληνικο' => 'κειμενο'}
      template = Template.new(fixture('ods/greek.ods'), tmp('output_greek.ods'))
      template.process binding
      expect(tmp('output_greek.ods')).to contain_in('content.xml', 'κειμενο')
    end

    it "loops and generates table rows" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('ods/table_rows.ods'), tmp('output_table_rows.ods'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_table_rows.ods')).to contain_in('content.xml', text)
      end
    end
  end

  describe 'PDF conversion' do
    it "converts an ODT document to PDF" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('odt/variables.odt'), tmp('output_variables.pdf'))
      template.process binding

      expect(tmp('output_variables.pdf')).to be_a_document
      expect(File.exist?(tmp('output_variables.odt'))).to be false
    end

    it "converts a DOCX document to PDF" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('docx/variables.docx'), tmp('output_variables_docx.pdf'))
      template.process binding

      expect(tmp('output_variables_docx.pdf')).to be_a_document
      expect(File.exist?(tmp('output_variables_docx.docx'))).to be false
    end

    it "raises an error when no converter is found" do
      @name = 'test'

      template = Template.new(fixture('odt/variables.odt'), tmp('output_pdf_none.pdf'))
      allow(template).to receive(:which).and_return(nil)

      expect { template.process binding }.to raise_error(RuntimeError, /No PDF converter found/)

      FileUtils.rm_f(tmp('output_pdf_none.odt'))
    end
  end

  describe XlsxProcessor do
    it "processes a document with simple variable substitution" do
      @name = 'Malcolm Reynolds'
      @title = 'captain'

      template = Template.new(fixture('xlsx/variables.xlsx'), tmp('output_variables.xlsx'))
      template.process binding

      expect(tmp('output_variables.xlsx')).to contain_in('xl/sharedStrings.xml', 'Malcolm Reynolds')
      expect(tmp('output_variables.xlsx')).to contain_in('xl/sharedStrings.xml', 'captain')
    end

    it "unrolls a simple for loop" do
      @crew = %w{'River', 'Jayne', 'Wash'}

      template = Template.new(fixture('xlsx/loop.xlsx'), tmp('output_loop.xlsx'))
      template.process binding
    end

    it "unrolls an advanced loop with tables" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('xlsx/loop_table.xlsx'), tmp('output_loop_table.xlsx'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_loop_table.xlsx')).to contain_in('xl/sharedStrings.xml', text)
      end
    end

    it "processes an advanced document" do
      PersonXlsx = Struct.new(:nome) unless defined?(PersonXlsx)
      @persons = [
        PersonXlsx.new('Malcolm'),
        PersonXlsx.new('River'),
        PersonXlsx.new('Jay')
      ]

      template = Template.new(fixture('xlsx/advanced.xlsx'), tmp('output_advanced.xlsx'))
      template.process binding

      ['Malcolm', 'River', 'Jay'].each do |text|
        expect(tmp('output_advanced.xlsx')).to contain_in('xl/sharedStrings.xml', text)
      end
    end

    it "processes a greek document" do
      @h = {'ελληνικο' => 'κειμενο'}
      template = Template.new(fixture('xlsx/greek.xlsx'), tmp('output_greek.xlsx'))
      template.process binding
      expect(tmp('output_greek.xlsx')).to contain_in('xl/sharedStrings.xml', 'κειμενο')
    end

    it "loops and generates table rows" do
      @ships = [Ship.new('Firefly', 'transport'), Ship.new('Colonial', 'battle')]

      template = Template.new(fixture('xlsx/table_rows.xlsx'), tmp('output_table_rows.xlsx'))
      template.process binding

      ['Firefly', 'transport', 'Colonial', 'battle'].each do |text|
        expect(tmp('output_table_rows.xlsx')).to contain_in('xl/sharedStrings.xml', text)
      end
    end
  end
end
