class String
  def escape_xml
    mgsub!([[/&/, '&amp;'], [/</, '&lt;'], [/>/, '&gt;']])
  end

  def convert_newlines
    format = Thread.current[:serenity_report_format]
    return self if format == :xlsx
    tag = format == :docx ? '<w:br/>' : '<text:line-break/>'
    gsub!("\n", tag)
    self
  end

  def mgsub!(key_value_pairs=[].freeze)
    regexp_fragments = key_value_pairs.collect { |k,v| k }
    gsub!(Regexp.union(*regexp_fragments)) do |match|
      key_value_pairs.detect{|k,v| k =~ match}[1]
    end
    self
  end
end
