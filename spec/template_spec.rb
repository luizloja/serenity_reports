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

      # Each person gets their own chart with their name and column values
      { 'Object 1' => ['Malcolm', '10.5', '20.3', '30.1'],
        'Object 2' => ['River',   '40.2', '50.7', '60.4'],
        'Object 3' => ['Jay',     '70.8', '80.9', '90.6'] }.each do |obj, values|
        values.each do |val|
          expect(tmp('output_advanced.odt')).to contain_in("#{obj}/content.xml", val)
        end
      end
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
end
