require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe String do
  it 'escapes <' do
    expect('1 < 2'.escape_xml).to eq('1 &lt; 2')
  end

  it 'escapes >' do
    expect('2 > 1'.escape_xml).to eq('2 &gt; 1')
  end

  it 'escapes &' do
    expect('1 & 2'.escape_xml).to eq('1 &amp; 2')
  end

  it 'escapes < > &' do
    expect('1 < 2 && 2 > 1'.escape_xml).to eq('1 &lt; 2 &amp;&amp; 2 &gt; 1')
  end

  it 'works on frozen strings with escape_xml' do
    frozen = 'frozen <value> & stuff'.freeze
    expect { frozen.escape_xml }.not_to raise_error
    expect(frozen.escape_xml).to eq('frozen &lt;value&gt; &amp; stuff')
  end

  it 'does not mutate the original string with escape_xml' do
    original = '1 < 2'
    original.escape_xml
    expect(original).to eq('1 < 2')
  end

  it 'works on frozen strings with convert_newlines' do
    frozen = "line1\nline2".freeze
    Thread.current[:serenity_report_format] = :odt
    expect { frozen.convert_newlines }.not_to raise_error
    expect(frozen.convert_newlines).to eq('line1<text:line-break/>line2')
  ensure
    Thread.current[:serenity_report_format] = nil
  end

  it 'does not mutate the original string with convert_newlines' do
    original = "line1\nline2"
    Thread.current[:serenity_report_format] = :odt
    original.convert_newlines
    expect(original).to eq("line1\nline2")
  ensure
    Thread.current[:serenity_report_format] = nil
  end

  it 'converts newlines to <w:br/> for docx format' do
    Thread.current[:serenity_report_format] = :docx
    expect("a\nb".convert_newlines).to eq('a<w:br/>b')
  ensure
    Thread.current[:serenity_report_format] = nil
  end

  it 'skips newline conversion for xlsx format' do
    Thread.current[:serenity_report_format] = :xlsx
    expect("a\nb".convert_newlines).to eq("a\nb")
  ensure
    Thread.current[:serenity_report_format] = nil
  end
end
