class String
  def escape_xml
    gsub(/[&<>]/, '&' => '&amp;', '<' => '&lt;', '>' => '&gt;')
  end

  def convert_newlines
    format = Thread.current[:serenity_report_format]
    return self if format == :xlsx
    tag = format == :docx ? '<w:br/>' : '<text:line-break/>'
    gsub("\n", tag)
  end
end
