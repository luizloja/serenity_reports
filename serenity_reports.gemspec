Gem::Specification.new do |s|
  s.name        = "serenity_reports"
  s.version     = "0.2.0"
  s.authors     = ["Tomas Kramar", "Luiz Loja"]
  s.email       = ["kramar.tomas@gmail.com", "luizloja@gmail.com"]
  s.homepage    = "https://github.com/luizloja/serenity_reports"
  s.summary     = "Parse ODT or DOCX file and substitutes placeholders like ERb."
  s.description = "Embedded ruby for OpenOffice/LibreOffice Text Document (.odt) and Word (.docx) files. You provide a template with ruby code in a special markup and the data, and SerenityReport generates the document. Very similar to .erb files."
  s.license     = "MIT"

  s.required_ruby_version = ">= 3.0"
  s.files = Dir["{lib}/**/*"] + ["LICENSE", "Rakefile", "README.md"]
  s.add_dependency "rubyzip", ">= 2.0"
  s.add_dependency "nokogiri", ">= 1.10"
  s.add_development_dependency "rspec", "~> 3.0"
end
